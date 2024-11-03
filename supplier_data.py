"""
Supplier data
"""

from dataclasses import dataclass, asdict
from datetime import datetime, timedelta
import json
import pandas as pd

@dataclass
class Supplier:
    """
    Supplier class
    """
    name: str
    email: str
    late_orders: dict
    preventive_orders: dict


class Suppliers:
    """
    Suppliers class
    """
    instances: dict[str, dict] = {}

    def __init__(self):
        self.email_data = pd.DataFrame()
        self.orders_data = pd.DataFrame()

    @classmethod
    def create(cls, name: str, email: str, late_orders: dict, preventive_orders: dict) -> None:
        """
        Create a new supplier instance
        """
        instance = Supplier(name, email, late_orders, preventive_orders)
        cls.instances[" ".join(instance.name.split(' ')
                               [:2])] = asdict(instance)

    def format_orders_data(self) -> None:
        """
        Format orders data with filters and datetime
        """
        mp_filter: list[str] = [
            'MATERIA-PRIMA',
            'MATERIA PRIMA INDUSTRIALIZAÇÃO',
            'MATERIAL DE USO E CONSUMO',
            'MATÉRIA PRIMA CABOS',
            'EMBALAGEM (MAT EMBALAGEM)'
        ]

        self.orders_data = self.orders_data[
            (self.orders_data['Situação'] != 'Envio pendente') &
            (self.orders_data['Nacionalidade'] == 'Brasil') &
            (self.orders_data['Rateio'].isin(mp_filter))
        ]

        self.orders_data['Data de entrega'] = pd.to_datetime(
            self.orders_data['Data de entrega'], format='%d/%m/%Y')

    def get_data_from_file(self, order_file_path: str, email_file_path: str) -> None:
        """
        Get suppliers data (name, email, orders) from files
        """
        today_date = datetime.now()

        self.email_data = pd.read_excel(
            email_file_path, usecols=["Nome", "Email"])

        self.orders_data = pd.read_excel(order_file_path, usecols=[
            "Neg.", "Data de entrega", "Fornecedor", "Cod.", "Material", "Faltam", "Nacionalidade",
            "Rateio", "Situação"])
        self.format_orders_data()

        total_late_orders: pd.DataFrame = self.orders_data[self.orders_data['Data de entrega']
                                                           < (today_date - timedelta(days=1))]
        total_preventive_orders: pd.DataFrame = self.orders_data[
            (self.orders_data['Data de entrega'] > today_date)
            & (self.orders_data['Data de entrega'] <= (today_date + timedelta(days=30)))]

        suppliers_names: list[str] = self.orders_data.loc[:, ['Fornecedor']].drop_duplicates(
            subset='Fornecedor', keep='first').squeeze().values

        for name in suppliers_names:
            email_data = self.email_data.loc[self.email_data['Nome'].str.strip() == name.strip(), [
                "Email"]]

            email: str = "; ".join(email_data.to_string(index=False, header=False).split(
                sep=',')).strip() if not email_data.empty else ""

            late_orders = total_late_orders.loc[total_late_orders['Fornecedor'] == name, [
                'Neg.', 'Data de entrega', 'Cod.', 'Material', 'Faltam']]
            late_orders['Data de entrega'] = late_orders['Data de entrega'].dt.strftime(
                "%d/%m/%Y")
            late_orders = late_orders.to_dict(orient='records')

            preventive_orders = total_preventive_orders.loc[
                total_preventive_orders['Fornecedor'] == name,
                    ['Neg.', 'Data de entrega', 'Cod.', 'Material', 'Faltam']]

            preventive_orders['Data de entrega'] = preventive_orders['Data de entrega'].dt.strftime(
                "%d/%m/%Y")
            preventive_orders = preventive_orders.to_dict(orient='records')

            if late_orders != [] or preventive_orders != []:
                Suppliers.create(name, email, late_orders, preventive_orders)

    def to_json(self):
        """
        Convert suppliers data to json
        """
        with open('supplier_data.json', 'w', encoding='utf-8') as file:
            json.dump(self.instances, file, indent=4, ensure_ascii=False)

        supplier: str = json.dumps(
            Suppliers.instances, indent=4, ensure_ascii=False)
        return json.loads(supplier)
