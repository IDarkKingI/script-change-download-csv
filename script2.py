import pandas as pd
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from datetime import datetime


def load_data(file_path):
    return pd.read_csv(file_path)

def display_csv_with_pandas(data):
    print(data.head())

def decimal_to_hours(decimal_time):
    hours = int(decimal_time)
    minutes = int((decimal_time - hours) * 60)
    return f"{hours} ч {minutes} мин"

def get_rate(employee, project):
    rate = employee_hourly_rates.get(employee, 0)
    if isinstance(rate, dict): 
        return rate.get(project, 0)
    return rate 

data = load_data("3.csv")

decimal_columns = [col for col in data.columns if "Total (Decimal)" in col and "Общая сумма" not in col and not col.startswith("Total")]
decimal_data = data[['Проект'] + decimal_columns]

decimal_data.columns = [col.replace(" Total (Decimal)", "") for col in decimal_data.columns]

melted_decimal_data = decimal_data.melt(id_vars=['Проект'], var_name='Сотрудник', value_name='Время')

grouped_decimal_data = melted_decimal_data.groupby(['Сотрудник', 'Проект']).sum().reset_index()

grouped_decimal_data['Нормальное время'] = grouped_decimal_data['Время'].apply(decimal_to_hours)

employee_hourly_rates = {
    'Akhmedov Daniyar': {'Назначенных проектов нет': 12, 'Передаче курса МБА в Тайвань': 12},
    'Alatortseva Larisa': 0,
    #'Akhmedov Daniyar': 12,
    'Artemchik Artem': 7,
    'Borsuk Anastasiya': 7.5,
    'Brezhneva Irina': 4.5,
    'Business Booster FB': 0,
    'Chekhlomin Sergey': 14.5,
    'Demchenko Olga': 17,
    'Dervoed Darya': 0,
    'Efremova Polina': 8.5,
    'Fadeev Georgiy': 15,
    'Fedyanina Darya': 5,
    'Frolov Aleksander': 0,
    'Galustyan Grisha': 6,
    'Gorshkova Darya': 5.5,
    'Gostroborodov Maksim': 6.5,
    'Gumerov Ruslan': 12,
    'Ilina Nadezhda': 32.5,
    'Karaseva Anastasiya': 0,
    'Khaev Dmitriy': 4.2,
    'Khoroshylova Olga': 6,
    'Kisilenko Olga': 6,
    'Kirichok Kateryna': 8.5,
    'Kotsolko Igor': 8.5,
    'Kritskaya Yuliya': 0,
    'Kudryavtsev Andrey': 0,
    'Kutovaya Kseniya': 0,
    'Kuznetsov Pavel': 0,
    'Kuznetsova Veronika': 0,
    'Lapin Daniel': 0,
    'Lavryk Natyaliya': 25,
    'Lidich Vitalii': 6,
    'Lutsenko Yuliya': 12,
    'Malyhina Svetlana': 9,
    'Merkulova Yuliya': 5.5,
    'Mishonin Nikolay': 0,
    'Mitin Alexey': 0,
    'Musienko Maria': 6.5,
    'Nikita Andreev': 0,
    'Roman Mosin': 0,
    'Salkov Andrey': 0,
    'Samadinova Elina': 7,
    'Shalomay Ekaterina': 15.1,
    'Shedko Valeriya': 6,
    'Sheremetiev Vitaly': 0,
    'Skakun Aleksandra': 6,
    'Smolina Mariya': 5,
    'Stepanchuk Anna': 15,
    'Telagys Kuanysh': 0,
    'Tkachenko Valeriya': 0,
    'Tsisinevich Veranika': 5,
    'Utkin Maksim': 5.5,
    'Uzdeeva Victoria': 12.5,
    'Zelianiuk Yelizaveta': 0,
    'Зара Берсунькаева': 0
}

grouped_decimal_data['Стоимость работы'] = grouped_decimal_data.apply(
    lambda row: row['Время'] * get_rate(row['Сотрудник'], row['Проект']), axis=1
)


total_time_and_costs = grouped_decimal_data.groupby('Сотрудник').agg(
    Общая_сумма_выплат=('Стоимость работы', 'sum'),
    Общее_время=('Время', 'sum')
).reset_index()

total_time_and_costs['Общее время (часы и минуты)'] = total_time_and_costs['Общее_время'].apply(decimal_to_hours)

final_data = grouped_decimal_data.merge(total_time_and_costs, on='Сотрудник')

filtered_data = final_data[final_data['Время'] > 0].copy()
filtered_data.rename(columns={
    'Общая_сумма_выплат': 'Общая сумма выплат',
    'Общее_время': 'Общее время',
    'Общее время (часы и минуты)': 'Общее время в часах и минутах'
}, inplace=True)


current_date = datetime.now().strftime("%Y-%m-%d")
report_name = f"Отчетный период - {current_date} - заработная плата.xlsx"
output_file_path = report_name
filtered_data.to_excel(output_file_path, index=False, engine='openpyxl')

print(f"Data saved to {output_file_path}")

SERVICE_ACCOUNT_FILE = 'for-zoom-api.json'
SCOPES = ['https://www.googleapis.com/auth/drive']

def upload_file_to_drive(filename, filepath, folder_id):
    credentials = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    service = build('drive', 'v3', credentials=credentials)
    file_metadata = {'name': filename, 'parents': [folder_id]}
    media = MediaFileUpload(filepath, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    try:
        # Добавление параметра supportsAllDrives
        file = service.files().create(body=file_metadata, media_body=media, fields='id', supportsAllDrives=True).execute()
        print('File ID: %s' % file.get('id'))
    except Exception as e:
        print(f"An error occurred: {e}")


folder_id = '12l2vTIusVndN3bBUEzZalLnCil4_9hZ-'
upload_file_to_drive(report_name, output_file_path, folder_id)