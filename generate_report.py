import pandas as pd

def generate_report():
    df = pd.read_excel("Result.xlsx")

    # Общая статистика
    stats = pd.DataFrame({
        "Метрика": ["Строк обработано", "Уникальные регионы"],
        "Значение": [
            len(df),
            df["region_id"].nunique()
        ]
    })

    # Неизвестные регионы
    unknown = df[df["status"] == "not_found"][["original_region"]].drop_duplicates()
    unknown.columns = ["Неизвестные регионы (нет в справочнике)"]

    # Аномалии
    anomalies = df[df["status"].isin(["empty", "country"])][
        ["source", "original_region", "status"]
    ]
    anomalies.columns = ["Источник", "Название региона", "Статус из JSON"]

    # Сохранение
    with pd.ExcelWriter("Report.xlsx") as writer:
        stats.to_excel(writer, sheet_name="Сводка", index=False)
        unknown.to_excel(writer, sheet_name="Неизвестные регионы", index=False)
        anomalies.to_excel(writer, sheet_name="Аномалии", index=False)

    print("Отчёт готов: Report.xlsx")


if __name__ == "__main__":
    generate_report()