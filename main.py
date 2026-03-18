import pandas as pd
from openai import OpenAI
import json

class LLMProcessor:
    def __init__(self):
        # Процессор для работы с локальной моделью из Ollama

        self.client = OpenAI(
            api_key="ollama",
            base_url="http://localhost:11434/v1"
        )

    def load_prompt(self, prompt_file="prompt.txt"):
       # Чтение промпта из файла

        with open(prompt_file, 'r', encoding='utf-8') as f:
            return f.read()

    def load_data(self, excel_file, regions_file):
        # Загрузка всех листов
        excel_data = pd.read_excel(excel_file, sheet_name=None, header=None)

        records = []
        for sheet_name, df in excel_data.items():
            # Проверка, что в листе достаточно строк
            if len(df) >= 5:
                # Данные берутся по фиксированным ячейкам:

                geo = str(df.iloc[2, 1]) if pd.notna(df.iloc[2, 1]) else ""
                segment = str(df.iloc[3, 1]) if pd.notna(df.iloc[3, 1]) else ""
                source = str(df.iloc[4, 1]) if pd.notna(df.iloc[4, 1]) else ""

                if source:  # если есть номер
                    records.append({
                        'source': source,
                        'region': geo,
                        'segment': segment,
                        'sheet': sheet_name
                    })

        # Загрузка справочника регионов
        regions_df = pd.read_excel(regions_file)
        regions_dict = dict(zip(regions_df.iloc[:, 0], regions_df.iloc[:, 1]))

        print(f"Загружено записей: {len(records)}")
        print(f"Загружено регионов: {len(regions_dict)}")

        return records, regions_dict

    def process(self, excel_file, regions_file, output_file, prompt_file="prompt.txt"):
        # Основной метод обработки

        print("Начало обработки")

        # Загрузка данных
        records, regions_dict = self.load_data(excel_file, regions_file)

        if not records:
            print("Нет данных для обработки")
            return

        # Подготовка и загрузка промпта
        prompt_template = self.load_prompt(prompt_file)
        regions_list = "\n".join([f"{r}: {regions_dict[r]}" for r in regions_dict])
        system_prompt = prompt_template.replace("{regions_list}", regions_list)

        # Данные для отправки
        user_data = []
        for r in records:
            user_data.append({
                'source': r['source'],
                'region': r['region'],
                'segment': r['segment']
            })

        print("Отправка данных в LLM...")

        # С
        try:
            response = self.client.chat.completions.create(
                model="deepseek-r1:8b", # модель deepseek, установленная на локальной Ollama
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": f"Обработай:\n{json.dumps(user_data, ensure_ascii=False, indent=2)}"}
                ],
                temperature=0.1, # 0.1 для более вероятных ответов
                max_tokens=4000 # 4000 токенов с запасом хватит для 10 листов
            )

            # Парсинг ответа из JSON
            result = json.loads(response.choices[0].message.content)

            # Сохранение
            df = pd.DataFrame(result if isinstance(result, list) else result.get('processed_records', []))
            df.to_excel(output_file, index=False)

            print(f"Результат сохранен в {output_file}")

            # Статистика
            if 'status' in df.columns:
                print("\n Статистика:")
                print(df['status'].value_counts())

        except Exception as e:
            print(f"Ошибка: {e}")


# Использование
def main():

    processor = LLMProcessor()
    processor.process(
        excel_file="Cleaned Data.xlsx",
        regions_file="Regions number.xlsm",
        output_file="Result.xlsx",
        prompt_file="prompt.txt"
    )


if __name__ == "__main__":
    main()