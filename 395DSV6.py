import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
from tkcalendar import DateEntry
import requests
from datetime import datetime, timedelta
import xml.etree.ElementTree as ET
from decimal import Decimal, ROUND_HALF_UP
import re
import json
import os

class KeyRateAPI:
    """Класс для работы с SOAP API ключевой ставки ЦБ РФ"""
    
    CACHE_FILE = "key_rates_cache.json"
    
    @classmethod
    def fetch_key_rates(cls):
        """
        Получает динамику ключевой ставки через SOAP API Банка России (KeyRate)
        Возвращает список (date, rate)
        """
        url = "http://www.cbr.ru/DailyInfoWebServ/DailyInfo.asmx"
        from_date = "2013-01-01"  # формат ISO
        to_date = datetime.today().strftime("%Y-%m-%d")
        soap_body = f"""
        <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
          <soap:Body>
            <KeyRate xmlns="http://web.cbr.ru/">
              <fromDate>{from_date}</fromDate>
              <ToDate>{to_date}</ToDate>
            </KeyRate>
          </soap:Body>
        </soap:Envelope>
        """
        headers = {
            "Content-Type": "text/xml; charset=utf-8",
            "SOAPAction": "http://web.cbr.ru/KeyRate"
        }
        try:
            response = requests.post(url, data=soap_body.encode('utf-8'), headers=headers, timeout=15)
            response.raise_for_status()
            root = ET.fromstring(response.content)
            ns = {'soap': 'http://schemas.xmlsoap.org/soap/envelope/', 'm': 'http://web.cbr.ru/'}
            result = root.find('.//m:KeyRateResult', ns)
            if result is None:
                print("Не найден KeyRateResult в ответе Банка России")
                return []
            # result.text может быть None, поэтому ищем первый дочерний элемент
            inner_xml = None
            for child in result:
                inner_xml = child
                break
            if inner_xml is None:
                print("Не найден вложенный XML с данными о ставках")
                return []
            # Найти контейнер <KeyRate>
            keyrate_container = None
            for elem in result.iter():
                if elem.tag.endswith('KeyRate'):
                    keyrate_container = elem
                    break
            if keyrate_container is None:
                print("Не найден контейнер <KeyRate>")
                return []
            rates = []
            for kr in keyrate_container.iter():
                if kr.tag.endswith('KR'):
                    date_str = None
                    rate_str = None
                    for child in kr:
                        if child.tag.endswith('DT'):
                            date_str = child.text
                        if child.tag.endswith('Rate'):
                            rate_str = child.text
                    if date_str and rate_str:
                        try:
                            date = datetime.strptime(date_str[:10], "%Y-%m-%d").date()
                            rate = Decimal(rate_str.replace(',', '.'))
                            rates.append((date, rate))
                        except Exception as ex:
                            print(f'Ошибка парсинга: {date_str} {rate_str} {ex}')
                            continue
            rates.sort()
            
            # Сохраняем в кэш при успешном получении
            cls.save_rates_to_cache(rates)
            return rates
        except Exception as e:
            print(f"Ошибка при обращении к сервису ЦБ РФ: {e}")
            # При ошибке сети пытаемся загрузить из кэша
            cached_rates = cls.load_rates_from_cache()
            if cached_rates:
                print("Используются кэшированные данные о ставках")
            return cached_rates
    
    @classmethod
    def save_rates_to_cache(cls, rates):
        """Сохраняет ставки в кэш-файл"""
        try:
            cache_data = {
                'last_update': datetime.now().isoformat(),
                'rates': [(date.isoformat(), float(rate)) for date, rate in rates]
            }
            with open(cls.CACHE_FILE, 'w', encoding='utf-8') as f:
                json.dump(cache_data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"Ошибка сохранения кэша: {e}")
    
    @classmethod
    def load_rates_from_cache(cls):
        """Загружает ставки из кэш-файла"""
        try:
            if not os.path.exists(cls.CACHE_FILE):
                return []
            
            with open(cls.CACHE_FILE, 'r', encoding='utf-8') as f:
                cache_data = json.load(f)
            
            rates = []
            for date_str, rate in cache_data['rates']:
                date = datetime.fromisoformat(date_str).date()
                rates.append((date, Decimal(str(rate))))
            
            rates.sort()
            return rates
        except Exception as e:
            print(f"Ошибка загрузки кэша: {e}")
            return []
    
    @classmethod
    def get_latest_rate_info(cls, rates):
        """Возвращает информацию о последней актуальной ставке"""
        if not rates:
            return None, None, None
        
        latest_date, latest_rate = rates[-1]
        return latest_date, latest_rate, len(rates)

    @staticmethod
    def get_rate_for_date(date, rates):
        """Поиск ставки по дате"""
        for i in range(len(rates)-1, -1, -1):
            if date >= rates[i][0]:
                return rates[i][1]
        return rates[0][1] if rates else Decimal('0')

class FinancialEvent:
    """Класс для хранения финансового события"""
    
    def __init__(self, date, amount, event_type):
        self.date = date
        self.amount = Decimal(str(amount))
        self.event_type = event_type  # 'debt' - долг, 'payment' - погашение, 'increase' - увеличение
    
    def __str__(self):
        type_str = {
            'debt': 'Начальный долг',
            'payment': 'Погашение', 
            'increase': 'Увеличение долга'
        }
        return f"{self.date.strftime('%d.%m.%Y')}: {type_str[self.event_type]} {self.amount}"

class Calculator395:
    """Калькулятор процентов по ст. 395 ГК РФ"""
    
    def __init__(self):
        self.events = []
        self.key_rates = []
        self.using_cached_data = False
        self.latest_rate_info = None
    
    def add_event(self, date, amount, event_type):
        """Добавление финансового события"""
        self.events.append(FinancialEvent(date, amount, event_type))
        self.events.sort(key=lambda x: x.date)
    
    def clear_events(self):
        """Очистка всех событий"""
        self.events.clear()
    
    def load_key_rates(self):
        """Загрузка ключевых ставок"""
        try:
            self.key_rates = KeyRateAPI.fetch_key_rates()
            if self.key_rates:
                self.latest_rate_info = KeyRateAPI.get_latest_rate_info(self.key_rates)
                return True
            return False
        except Exception as e:
            # Пытаемся загрузить из кэша при ошибке
            cached_rates = KeyRateAPI.load_rates_from_cache()
            if cached_rates:
                self.key_rates = cached_rates
                self.latest_rate_info = KeyRateAPI.get_latest_rate_info(self.key_rates)
                self.using_cached_data = True
                return True
            return False
    
    def days_in_year(self, date):
        """Проверка високосного года и получение количества дней в году"""
        year = date.year
        return 366 if (year % 4 == 0 and (year % 100 != 0 or year % 400 == 0)) else 365
    
    def get_rate_for_date(self, date):
        """Получение ставки для указанной даты"""
        return KeyRateAPI.get_rate_for_date(date, self.key_rates)
    
    def calculate_interests(self, calc_date):
        """Основной расчет процентов - показываем только периоды между событиями"""
        if not self.events:
            return [], Decimal('0'), Decimal('0')
        
        if not self.key_rates:
            if not self.load_key_rates():
                messagebox.showerror("Ошибка", "Не удалось загрузить ключевые ставки")
                return [], Decimal('0'), Decimal('0')
        
        # Сортируем события
        self.events.sort(key=lambda x: x.date)
        
        periods = []
        current_debt = Decimal('0')
        total_interest = Decimal('0')
        
        # Проходим по всем событиям и создаем периоды между ними
        for i in range(len(self.events)):
            current_event = self.events[i]
            
            # Определяем начало периода
            # Для первого события (начала расчета) - включая дату события
            # Для промежуточных событий - со следующего дня
            if i == 0:
                start_date = current_event.date  # Начало расчета - включая дату
            else:
                start_date = current_event.date + timedelta(days=1)  # Промежуточные - со следующего дня
            
            # Определяем конец периода
            if i < len(self.events) - 1:
                # Для промежуточных периодов - до даты следующего события включительно
                end_date = self.events[i + 1].date
            else:
                # Для последнего периода - до даты расчета включительно
                end_date = calc_date
            
            if end_date < start_date:
                continue
            
            # Применяем текущее событие к долгу
            event_description = ""
            event_amount = Decimal('0')
            
            if current_event.event_type == 'debt':
                current_debt = current_event.amount
                event_description = "Начальный долг"
                event_amount = current_event.amount
            elif current_event.event_type == 'payment':
                payment_amount = min(current_event.amount, current_debt)
                current_debt -= payment_amount
                event_description = "Погашение"
                event_amount = -payment_amount
            elif current_event.event_type == 'increase':
                current_debt += current_event.amount
                event_description = "Увеличение долга"
                event_amount = current_event.amount
            
            # Пропускаем период если долг нулевой
            if current_debt == Decimal('0'):
                continue
            
            # Разбиваем период на подпериоды по изменениям ставки и годам
            sub_periods = self.split_period_by_rates_and_years(start_date, end_date, current_debt)
            
            for sub_period in sub_periods:
                # Событие отображаем только для первого подпериода
                if sub_period['start_date'] == start_date:
                    sub_period['event'] = event_description
                    sub_period['event_amount'] = event_amount
                else:
                    sub_period['event'] = ""
                    sub_period['event_amount'] = Decimal('0')
                
                periods.append(sub_period)
                total_interest += sub_period['interest']
        
        return periods, total_interest, current_debt
    
    def split_period_by_rates_and_years(self, start_date, end_date, debt):
        """Разбивает период на подпериоды по изменениям ставки и границам годов"""
        sub_periods = []
        current_date = start_date
        
        while current_date <= end_date:
            # Определяем конец текущего подпериода
            period_end = end_date
            
            # Проверяем изменение ставки
            current_rate = self.get_rate_for_date(current_date)
            next_date = current_date + timedelta(days=1)
            while next_date <= end_date:
                next_rate = self.get_rate_for_date(next_date)
                if next_rate != current_rate:
                    period_end = next_date - timedelta(days=1)
                    break
                next_date += timedelta(days=1)
            
            # Проверяем границу года
            current_year = current_date.year
            year_end = datetime(current_year, 12, 31).date()
            if year_end < period_end and year_end >= current_date:
                period_end = year_end
            
            # Убеждаемся, что период не выходит за границы
            period_end = min(period_end, end_date)
            
            # Расчет дней и процентов
            days = (period_end - current_date).days + 1
            year_days = self.days_in_year(current_date)
            interest = (debt * current_rate * Decimal(days) / 
                       Decimal('100') / Decimal(year_days))
            interest = interest.quantize(Decimal('0.01'), ROUND_HALF_UP)
            
            sub_periods.append({
                'start_date': current_date,
                'end_date': period_end,
                'debt': debt,
                'days': days,
                'rate': current_rate,
                'interest': interest,
                'year_days': year_days
            })
            
            current_date = period_end + timedelta(days=1)
        
        return sub_periods

class InterestCalculator(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("395 ГК РФ: Калькулятор процентов")
        self.geometry("1000x700")
        self.resizable(True, True)
        self.calculator = Calculator395()
        self.event_rows = []
        self.max_events = 100
        
        self.create_widgets()
        self.fetch_rates()
        
        # Назначаем обработчики клавиш
        self.bind('<Return>', lambda event: self.calculate())
        self.bind('<Tab>', self.on_tab)

    def on_tab(self, event):
        """Обработка клавиши Tab для переключения между элементами"""
        event.widget.tk_focusNext().focus()
        return "break"

    def create_widgets(self):
        # Дата расчета с календарём
        date_frame = tk.Frame(self)
        date_frame.pack(padx=10, pady=5, fill='x')
        tk.Label(date_frame, text="Дата расчета (по умолчанию сегодня):").pack(side='left')
        self.calc_date_var = tk.StringVar()
        self.calc_date_var.set(datetime.today().strftime('%d.%m.%Y'))
        self.calc_date_entry = DateEntry(date_frame, textvariable=self.calc_date_var, 
                                       date_pattern='dd.mm.yyyy', width=15)
        self.calc_date_entry.pack(side='left', padx=5)
        
        # Назначаем Tab для календаря
        self.calc_date_entry.bind('<Tab>', self.on_tab)

        # Информация о ставках
        self.info_frame = tk.Frame(self, bg='lightgray', relief='sunken', bd=1)
        self.info_frame.pack(padx=10, pady=5, fill='x')
        self.info_label = tk.Label(self.info_frame, bg='lightgray', wraplength=980, justify='left')
        self.info_label.pack(padx=5, pady=5, fill='x')

        # События
        self.events_frame = tk.LabelFrame(self, text="Финансовые события (до 100)")
        self.events_frame.pack(padx=10, pady=10, fill='both', expand=True)
        
        self.events_canvas = tk.Canvas(self.events_frame, height=200)
        self.events_canvas.pack(side='left', fill='both', expand=True)
        
        self.scrollbar = tk.Scrollbar(self.events_frame, orient='vertical', 
                                    command=self.events_canvas.yview)
        self.scrollbar.pack(side='right', fill='y')
        
        self.events_canvas.configure(yscrollcommand=self.scrollbar.set)
        self.events_inner = tk.Frame(self.events_canvas)
        self.events_canvas.create_window((0,0), window=self.events_inner, anchor='nw')
        
        self.events_inner.bind('<Configure>', 
                             lambda e: self.events_canvas.configure(
                                 scrollregion=self.events_canvas.bbox('all')))

        # Заголовки
        headers = ["#", "Дата (ДД.ММ.ГГГГ)", "Тип", "Сумма"]
        for i, h in enumerate(headers):
            tk.Label(self.events_inner, text=h, font=('Arial', 10, 'bold')).grid(
                row=0, column=i, padx=2, pady=2)

        # Добавляем первую строку с начальным долгом
        self.add_event_row("Начальный долг", readonly_type=True)

        self.add_event_btn = tk.Button(self, text="Добавить событие", 
                                     command=self.add_event_row)
        self.add_event_btn.pack(pady=5)
        
        # Назначаем Tab для кнопки добавления
        self.add_event_btn.bind('<Tab>', self.on_tab)

        self.calc_btn = tk.Button(self, text="Рассчитать (Enter)", 
                                command=self.calculate)
        self.calc_btn.pack(pady=5)
        
        # Назначаем Tab для кнопки расчета
        self.calc_btn.bind('<Tab>', self.on_tab)

        self.result_area = scrolledtext.ScrolledText(self, width=120, height=20, 
                                                   font=("Consolas", 9))
        self.result_area.pack(padx=10, pady=10, fill='both', expand=True)
        
        # Назначаем Tab для области результатов
        self.result_area.bind('<Tab>', self.on_tab)

    def update_rate_info(self):
        """Обновляет информацию о ставках в сером поле"""
        if self.calculator.latest_rate_info:
            latest_date, latest_rate, rates_count = self.calculator.latest_rate_info
            if latest_date and latest_rate:
                info_text = f"По данным ЦБ РФ на {latest_date.strftime('%d.%m.%Y')} установлена ключевая ставка {latest_rate}%"
                
                if self.calculator.using_cached_data:
                    info_text += f"\nВНИМАНИЕ: Соединение с ЦБ РФ недоступно. Используются кэшированные данные от {latest_date.strftime('%d.%m.%Y')}"
                else:
                    info_text += f"\nЗагружено {rates_count} записей о ключевой ставке с 2013 года"
                
                self.info_label.config(text=info_text)
        else:
            self.info_label.config(text="Не удалось загрузить данные о ключевых ставках")

    def fetch_rates(self):
        self.result_area.insert('end', "Загрузка ключевых ставок ЦБ РФ...\n")
        self.update()
        try:
            if self.calculator.load_key_rates():
                self.update_rate_info()
                if self.calculator.using_cached_data:
                    self.result_area.insert('end', "Используются кэшированные данные о ставках (отсутствует интернет-соединение)\n")
                self.result_area.insert('end', f"Загружено {len(self.calculator.key_rates)} ставок.\n\n")
            else:
                self.result_area.insert('end', "Ошибка загрузки ставок\n")
                self.info_label.config(text="Ошибка загрузки данных о ключевых ставках")
        except Exception as e:
            self.result_area.insert('end', f"Ошибка загрузки ставок: {e}\n")
            self.info_label.config(text=f"Ошибка загрузки данных: {e}")

    def add_event_row(self, default_type="Погашение", readonly_type=False):
        if len(self.event_rows) >= self.max_events:
            messagebox.showwarning("Лимит", "Достигнут лимит событий (100).")
            return
            
        row = len(self.event_rows) + 1
        
        # Номер строки
        idx_lbl = tk.Label(self.events_inner, text=str(row))
        idx_lbl.grid(row=row, column=0, padx=2, pady=2)
        
        # Дата
        date_v = tk.StringVar()
        date_entry = DateEntry(self.events_inner, textvariable=date_v, 
                             date_pattern='dd.mm.yyyy', width=13)
        date_entry.grid(row=row, column=1, padx=2, pady=2)
        
        # Назначаем Tab для поля даты
        date_entry.bind('<Tab>', self.on_tab)
        
        # Тип события
        type_v = tk.StringVar(value=default_type)
        # Для первой строки доступен "Начальный долг", для остальных - только "Погашение" и "Увеличение долга"
        if row == 1:
            type_values = ["Погашение", "Увеличение долга", "Начальный долг"]
        else:
            type_values = ["Погашение", "Увеличение долга"]
            
        type_cb = ttk.Combobox(self.events_inner, textvariable=type_v, 
                              values=type_values, width=15, 
                              state='readonly' if readonly_type else 'readonly')
        type_cb.grid(row=row, column=2, padx=2, pady=2)
        
        # Назначаем Tab для combobox
        type_cb.bind('<Tab>', self.on_tab)
        
        # Сумма
        amount_v = tk.StringVar()
        amount_entry = tk.Entry(self.events_inner, textvariable=amount_v, width=15)
        amount_entry.grid(row=row, column=3, padx=2, pady=2)
        
        # Назначаем Tab для поля суммы и Enter для расчета
        amount_entry.bind('<Tab>', self.on_tab)
        amount_entry.bind('<Return>', lambda event: self.calculate())
        
        # Кнопка удаления (кроме первой строки)
        def remove_event():
            for widget in self.events_inner.grid_slaves(row=row):
                widget.destroy()
            self.event_rows = [r for r in self.event_rows if r[0] != date_v]
            self.renumber_rows()
            
        if row > 1:
            btn = tk.Button(self.events_inner, text="✖", command=remove_event, 
                          fg="red", width=2)
            btn.grid(row=row, column=4, padx=2, pady=2)
            
            # Назначаем Tab для кнопки удаления
            btn.bind('<Tab>', self.on_tab)
            
        self.event_rows.append((date_v, type_v, amount_v))
        self.events_canvas.yview_moveto(1)

    def renumber_rows(self):
        """Перенумеровывает строки после удаления"""
        for i, (date_v, type_v, amount_v) in enumerate(self.event_rows):
            for widget in self.events_inner.grid_slaves(row=i+1, column=0):
                if isinstance(widget, tk.Label):
                    widget.config(text=str(i+1))

    def get_events(self):
        """Получение событий из интерфейса"""
        events = []
        for date_v, type_v, amount_v in self.event_rows:
            date_str = date_v.get().strip()
            type_str = type_v.get().strip()
            amount_str = amount_v.get().strip()
            
            if not date_str or not type_str or not amount_str:
                continue
                
            try:
                date = datetime.strptime(date_str, "%d.%m.%Y").date()
                amount = Decimal(amount_str.replace(',', '.'))
                
                # Преобразование типа события
                event_type_map = {
                    'Начальный долг': 'debt',
                    'Погашение': 'payment', 
                    'Увеличение долга': 'increase'
                }
                
                events.append({
                    'date': date,
                    'type': event_type_map.get(type_str, 'payment'),
                    'amount': amount
                })
                
            except Exception as e:
                print(f"Ошибка парсинга события: {e}")
                continue
                
        return events

    def calculate(self):
        """Основной расчет"""
        self.result_area.delete('1.0', 'end')
        
        # Получаем дату расчета
        try:
            calc_date = datetime.strptime(self.calc_date_var.get(), "%d.%m.%Y").date()
        except:
            calc_date = datetime.today().date()
            self.calc_date_var.set(calc_date.strftime('%d.%m.%Y'))
        
        # Получаем события
        events_data = self.get_events()
        if not events_data:
            self.result_area.insert('end', "Нет корректных событий для расчета.\n")
            return
            
        # Очищаем и добавляем события в калькулятор
        self.calculator.clear_events()
        for event in events_data:
            self.calculator.add_event(event['date'], event['amount'], event['type'])
        
        # Выполняем расчет
        try:
            periods, total_interest, final_debt = self.calculator.calculate_interests(calc_date)
            self.display_results(periods, total_interest, final_debt)
        except Exception as e:
            self.result_area.insert('end', f"Ошибка расчета: {e}\n")

    def display_results(self, periods, total_interest, final_debt):
        """Отображение результатов расчета"""
        # Заголовок таблицы
        header = f"{'Долг':<13}{'Период с':<12}{'по':<12}{'дни':<5}{'Ставка':<7}{'Проценты':<12}{'Событие':<14}{'Сумма':<10}\n"
        self.result_area.insert('end', header)
        self.result_area.insert('end', '-' * 85 + '\n')
        
        # Данные периодов
        for period in periods:
            event_type_map = {
                'Начальный долг': 'Начальный долг',
                'Погашение': 'Погашение',
                'Увеличение долга': 'Увеличение',
                'Конец периода': 'Конец периода',
                '': ''
            }
            
            event_str = event_type_map.get(period['event'], period['event'])
            amount_str = f"{period['event_amount']:.2f}" if period['event_amount'] != 0 else ""
            
            line = (f"{period['debt']:<13.2f}"
                   f"{period['start_date'].strftime('%d.%m.%Y'):<12}"
                   f"{period['end_date'].strftime('%d.%m.%Y'):<12}"
                   f"{period['days']:<5}"
                   f"{period['rate']:<7.2f}"
                   f"{period['interest']:<12.2f}"
                   f"{event_str:<14}"
                   f"{amount_str:<10}\n")
            
            self.result_area.insert('end', line)
        
        # Итоги
        self.result_area.insert('end', '-' * 85 + '\n')
        self.result_area.insert('end', f"{'Итого процентов:':<60}{total_interest:.2f}\n")
        self.result_area.insert('end', f"Остаток основного долга: {final_debt:.2f}\n")

def main():
    """Запуск приложения"""
    app = InterestCalculator()
    app.mainloop()

if __name__ == "__main__":
    main()
