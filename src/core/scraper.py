import io
import os
import json
import asyncio
import logging
import pandas as pd
from bs4 import BeautifulSoup
from datetime import datetime
from src.core.session_manager import SessionManager


class ReportScraper:
    def __init__(self, session_manager: SessionManager):
        self.session_manager = session_manager
        self.base_url = session_manager.base_url
        self.today = datetime.now().date()
        self.today_str = self.today.strftime("%Y-%m-%d")

    async def get_all_reports(
        self, init_date: str, end_date: str, download_suppliers: bool = True
    ):
        """
        Fetches both suppliers and pending orders reports concurrently,
        merges them, saves the outputs (JSON/Excel), and returns the unified data.
        """
        logging.info(
            f"Starting parallel fetch for reports from {init_date} to {end_date}. "
            f"Download suppliers: {download_suppliers}"
        )

        pending_orders_task = asyncio.create_task(
            self.get_pending_orders_report(init_date, end_date, return_json=False)
        )

        if download_suppliers:
            suppliers_task = asyncio.create_task(
                self.get_suppliers_report(return_json=False)
            )
            suppliers, pending_orders = await asyncio.gather(
                suppliers_task, pending_orders_task
            )
        else:
            # Use static local file
            suppliers_path = os.path.join("tmp", "fornecedores.json")
            if os.path.exists(suppliers_path):
                logging.info(f"Using local suppliers list from: {suppliers_path}")
                with open(suppliers_path, "r", encoding="utf-8") as f:
                    suppliers = json.load(f)
            else:
                logging.warning(
                    f"Local suppliers list not found at {suppliers_path}. Fetching anyway."
                )
                suppliers = await self.get_suppliers_report(return_json=False)

            pending_orders = await pending_orders_task

        logging.info("Both reports successfully fetched. Starting unification process.")
        return self._unify_reports(suppliers, pending_orders)

    async def get_suppliers_report(
        self, grupo_id: str = "821573", return_json: bool = False
    ):
        """
        Retrieves the suppliers report as a CSV binary, parses it to a dictionary list,
        and optionally saves the raw JSON file.
        """
        url = f"{self.base_url}/relatorio/pessoa/exportarPessoa/"
        params = {
            "relatorioPessoas[grupoId]": grupo_id,
            "relatorioPessoas[pessoaAtiva]": "1",
            "relatorioPessoas[tipoPessoa]": "Todas",
        }
        try:
            logging.info(f"Requesting suppliers report for group ID: {grupo_id}")
            async with self.session_manager.session.get(url, params=params) as response:
                response.raise_for_status()
                content = await response.read()
                return await self._parse_suppliers_report_to_json(
                    content, return_json=return_json
                )
        except Exception as e:
            logging.error(f"Failed to fetch suppliers report: {e}")
            return []

    async def _parse_suppliers_report_to_json(
        self, content: bytes, return_json: bool = False
    ):
        """
        Parses raw CSV bytes into a filtered DataFrame and returns it as a list of dicts.
        """
        with io.BytesIO(content) as buffer:
            df = pd.read_csv(buffer, sep=";", encoding="latin-1", on_bad_lines="skip")

        if df.empty:
            logging.warning("Suppliers report is empty.")
            return []

        # Clean column names
        df.columns = [c.strip().replace('"', "") for c in df.columns]

        if "CPF/CNPJ" in df.columns:
            df.rename(columns={"CPF/CNPJ": "CNPJ"}, inplace=True)

        # Handle missing emails
        if "Email" in df.columns:
            df["Email"] = df["Email"].fillna("-")
            df.loc[df["Email"] == "", "Email"] = "-"

        # Strip strings and remove quotes
        df = df.apply(
            lambda x: x.str.strip().str.replace('"', "") if x.dtype == "object" else x
        )

        target_columns = ["Nome", "CNPJ", "Email"]
        existing_columns = [c for c in target_columns if c in df.columns]
        df_filtered = df[existing_columns].copy()

        records = df_filtered.to_dict(orient="records")

        if not return_json:
            json_str = json.dumps(records, ensure_ascii=False, indent=4)
            json_str = json_str.replace("\\/", "/")
            os.makedirs("tmp", exist_ok=True)
            with open("tmp/fornecedores.json", "w", encoding="utf-8") as f:
                f.write(json_str)
            logging.info("Raw suppliers JSON file generated successfully in tmp/.")

        return records

    async def get_pending_orders_report(
        self, init_date: str, end_date: str, return_json: bool = False
    ):
        """
        Retrieves the pending orders report as HTML, extracts the table,
        and optionally saves the raw JSON file.
        """
        url = f"{self.base_url}/relatorio/compra/renderGridExportacaoEntregasPendentes/"
        params = {
            "RelatorioEntregasPendentes[referencia]": "ENT",
            "RelatorioEntregasPendentes[previsaoInicio]": init_date,
            "RelatorioEntregasPendentes[previsaoFim]": end_date,
            "RelatorioEntregasPendentes[fornecedorId]": "",
            "RelatorioEntregasPendentes[tipoFrete]": "9",
        }
        try:
            logging.info(f"Requesting pending orders from {init_date} to {end_date}")
            async with self.session_manager.session.get(url, params=params) as response:
                response.raise_for_status()
                html_content = await response.text()
                return await self._parse_pending_orders_report_to_json(
                    html_content, return_json=return_json
                )
        except Exception as e:
            logging.error(f"Failed to fetch pending orders report: {e}")
            return []

    async def _parse_pending_orders_report_to_json(
        self, html_content: str, return_json: bool = False
    ):
        """
        Parses HTML content, extracts table rows, applies business filters, and returns a list of dicts.
        """
        soup = BeautifulSoup(html_content, "html.parser")
        table = soup.find("table")

        if not table:
            logging.warning("No HTML table found in the pending orders response.")
            return []

        rows = []
        headers = [th.get_text(strip=True) for th in table.find_all("th")]
        for tr in table.find_all("tr"):
            cells = [td.get_text(strip=True) for td in tr.find_all("td")]
            if cells:
                rows.append(cells)

        if not rows:
            logging.info("Pending orders table is empty for the selected period.")
            return []

        df = pd.DataFrame(rows, columns=headers if headers else None)
        df = self._filter_pending_orders(df)

        # Keep only essential columns
        df = df[["Neg.", "Data de entrega", "Fornecedor", "Cod.", "Material", "Faltam"]]
        records = df.to_dict(orient="records")

        if not return_json:
            json_data = json.dumps(records, ensure_ascii=False, indent=4)
            json_data = json_data.replace("\\/", "/")
            os.makedirs("tmp", exist_ok=True)
            with open("tmp/entregas_pendentes.json", "w", encoding="utf-8") as f:
                f.write(json_data)
            logging.info(
                f"Raw pending orders JSON generated ({len(df)} records) in tmp/."
            )

        return records

    def _filter_pending_orders(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Applies business rules to filter out irrelevant orders (e.g., non-Brazilian or not raw materials).
        """
        mp_filter: list[str] = [
            "MATERIA-PRIMA",
            "MATERIA PRIMA INDUSTRIALIZAÇÃO",
            "MATERIAL DE USO E CONSUMO",
            "MATÉRIA PRIMA CABOS",
            "EMBALAGEM (MAT EMBALAGEM)",
        ]
        df = df[
            (df["Situação"] != "Envio pendente")
            & (df["Nacionalidade"] == "Brasil")
            & (df["Rateio"].isin(mp_filter))
        ]
        return df

    def _unify_reports(self, suppliers, pending_orders):
        """
        Orchestrator method for combining suppliers and pending orders into a structured payload.
        """
        logging.info("Mapping suppliers data...")
        supplier_map, name_to_cnpj = self._map_suppliers(suppliers)

        logging.info("Classifying orders by status (late/future)...")
        unified_data, without_email_data = self._classify_orders(
            pending_orders, supplier_map, name_to_cnpj
        )

        self._save_json(unified_data)
        self._save_excel(unified_data)
        self._save_suppliers_without_email_excel(without_email_data)

        return unified_data

    def _map_suppliers(self, suppliers):
        """
        Normalizes supplier data and builds indexes for quick lookup.
        """
        supplier_map = {}
        name_to_cnpj = {}

        for s in suppliers:
            cnpj = str(s.get("CNPJ", "-")).strip()
            name = str(s.get("Nome", "")).strip()
            if not name:
                continue

            supplier_map[cnpj] = {
                "supplier_name": name,
                "cnpj": cnpj,
                "email": self._clean_email(s.get("Email")),
                "late_orders": [],
                "future_orders": [],
            }
            name_to_cnpj[name] = cnpj

        return supplier_map, name_to_cnpj

    def _clean_email(self, raw_email):
        """
        Cleans email strings (removes mailto:, standardizes separators, trims spaces).
        """
        if not raw_email or str(raw_email).lower() == "nan":
            return "-"

        email = str(raw_email).lower().replace("mailto:", "").replace(",", ";")
        parts = [e.strip() for e in email.split(";") if e.strip()]
        return "; ".join(parts)

    def _classify_orders(self, pending_orders, supplier_map, name_to_cnpj):
        """
        Distributes pending orders to their respective suppliers and calculates delay days.
        """
        for order in pending_orders:
            target_cnpj = name_to_cnpj.get(str(order.get("Fornecedor", "")).strip())
            if not target_cnpj:
                continue

            data_entrega = order.get("Data de entrega")
            dt_obj = self._parse_date(data_entrega)

            if not dt_obj:
                continue

            # Calculate days delayed
            dias_atraso = (self.today - dt_obj).days if dt_obj < self.today else 0

            order_info = {
                "Neg.": order.get("Neg."),
                "Data de entrega": data_entrega,
                "Cod.": order.get("Cod."),
                "Material": order.get("Material"),
                "Faltam": order.get("Faltam"),
                "Dias de Atraso": dias_atraso,
            }

            key = "late_orders" if dt_obj < self.today else "future_orders"
            supplier_map[target_cnpj][key].append(order_info)

        # Separate suppliers with orders into two groups: with valid email and without
        valid_suppliers = []
        without_email_suppliers = []

        for v in supplier_map.values():
            if v["late_orders"] or v["future_orders"]:
                if v["email"] in ["-", "NaN", None, ""]:
                    without_email_suppliers.append(v)
                else:
                    valid_suppliers.append(v)

        return valid_suppliers, without_email_suppliers

    def _parse_date(self, date_str):
        """
        Safely converts a string date (YYYY-MM-DD HH:MM:SS) to a datetime.date object.
        """
        try:
            if not date_str:
                return None
            return datetime.strptime(str(date_str).split(" ")[0], "%Y-%m-%d").date()
        except Exception:
            return None

    def _save_json(self, data):
        """
        Persists the unified data as a JSON file in the tmp/dados directory.
        """
        path = "tmp/dados"
        os.makedirs(path, exist_ok=True)
        file_path = os.path.join(path, f"{self.today_str}_unified_report.json")

        with open(file_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=4)
        logging.info(f"Unified JSON successfully saved to: {file_path}")

    def _save_excel(self, data):
        """
        Flattens the structured data and exports it as an Excel file in the tmp/relatorios directory.
        """
        path = "tmp/relatorios"
        os.makedirs(path, exist_ok=True)

        rows = []
        for s in data:
            for order in s["late_orders"] + s["future_orders"]:
                rows.append(
                    {
                        "Fornecedor": s["supplier_name"],
                        "CNPJ": s["cnpj"],
                        "Email": s["email"],
                        "Status": (
                            "Atrasado" if order["Dias de Atraso"] > 0 else "No Prazo"
                        ),
                        **order,
                    }
                )

        if rows:
            df = pd.DataFrame(rows)
            # Ensure BR date format in Excel
            if "Data de entrega" in df.columns:
                df["Data de entrega"] = pd.to_datetime(
                    df["Data de entrega"]
                ).dt.strftime("%d/%m/%Y")

            file_path = os.path.join(path, f"{self.today_str}_relatorio_compras.xlsx")
            df.to_excel(file_path, index=False)
            logging.info(f"Unified Excel report successfully saved to: {file_path}")

    def _save_suppliers_without_email_excel(self, data):
        """
        Exports suppliers that have orders but no email to a separate Excel file.
        """
        if not data:
            return

        path = "tmp/relatorios"
        os.makedirs(path, exist_ok=True)

        rows = []
        for s in data:
            for order in s["late_orders"] + s["future_orders"]:
                rows.append(
                    {
                        "Fornecedor": s["supplier_name"],
                        "CNPJ": s["cnpj"],
                        "Email": "PENDENTE",
                        "Status": (
                            "Atrasado" if order["Dias de Atraso"] > 0 else "No Prazo"
                        ),
                        **order,
                    }
                )

        if rows:
            df = pd.DataFrame(rows)
            # Ensure BR date format in Excel
            if "Data de entrega" in df.columns:
                df["Data de entrega"] = pd.to_datetime(
                    df["Data de entrega"]
                ).dt.strftime("%d/%m/%Y")

            file_path = os.path.join(
                path, f"{self.today_str}_fornecedores_sem_email.xlsx"
            )
            df.to_excel(file_path, index=False)
            logging.info(f"Report for suppliers without email saved to: {file_path}")
