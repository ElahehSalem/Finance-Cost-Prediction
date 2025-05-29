import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
from dateutil.relativedelta import relativedelta
import jdatetime


class LoanCalculatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("محاسبه گر هزینه های فاینانس")
        self.project_name = tk.StringVar()

        self.input_frame = ttk.LabelFrame(root, text="ورودی‌ها")
        self.input_frame.pack(padx=10, pady=10, fill="x")

        self.result_frame = ttk.LabelFrame(root, text="نتایج")
        self.result_frame.pack(padx=10, pady=10, fill="both", expand=True)

        self.create_input_widgets()
        self.create_result_table()

        self.style = ttk.Style()
        self.style.configure("Primary.TButton",
                             foreground="white",
                             background="#007bff",
                             font=('Tahoma', 10, 'bold'))

    def create_input_widgets(self):
        ttk.Label(self.input_frame, text="نام پروژه:").grid(row=0, column=0, padx=5, pady=2, sticky="w")
        ttk.Entry(self.input_frame, textvariable=self.project_name, width=20).grid(row=0, column=1, padx=5, pady=2,
                                                                                   sticky="e")

        widgets = [
            ("مبلغ قرارداد:", ttk.Entry, 1),
            ("میزان فاینانس (%):", ttk.Entry, 2, "85"),
            ("نرخ بهره سالانه (%):", ttk.Entry, 3),
            ("نرخ بیمه تسهیلات (%):", ttk.Entry, 4, "0"),
            ("تعداد پرداخت در سال:", ttk.Combobox, 5, [1, 2, 3, 4, 6, 12]),
            ("مدت اجرای طرح (سال):", ttk.Spinbox, 6, (1, 30)),
            ("تعداد سالهای بازپرداخت:", ttk.Spinbox, 7, (1, 30)),
            ("سال شروع:", ttk.Spinbox, 8, (2000, 2100)),
            ("ماه شروع:", ttk.Spinbox, 9, (1, 12)),
            ("دوره تنفس پس از اجرا (سال):", ttk.Spinbox, 10, (0, 10)),
            ("دوره تنفس پس از اجرا (ماه):", ttk.Spinbox, 11, (0, 11)),
            ("نرخ هزینه مدیریت (%):", ttk.Entry, 12, "0"),
            ("نرخ هزینه تعهد (%):", ttk.Entry, 13, "0"),
            ("تعداد پرداخت تعهد در سال:", ttk.Combobox, 14, [1, 2, 3, 4, 6, 12])
        ]

        for i, widget in enumerate(widgets, start=1):
            label = ttk.Label(self.input_frame, text=widget[0])
            label.grid(row=i, column=0, padx=5, pady=2, sticky="w")

            if widget[1] == ttk.Entry:
                entry = ttk.Entry(self.input_frame, width=15)
                if len(widget) > 3:
                    entry.insert(0, widget[3])
            elif widget[1] == ttk.Combobox:
                entry = ttk.Combobox(self.input_frame, values=widget[3], width=13)
                entry.current(0)
            elif widget[1] == ttk.Spinbox:
                entry = ttk.Spinbox(self.input_frame,
                                    from_=widget[3][0],
                                    to=widget[3][1],
                                    width=10)

            entry.grid(row=i, column=1, padx=5, pady=2, sticky="e")
            setattr(self, f"input_{i}", entry)

        ttk.Button(self.input_frame,
                   text="محاسبه",
                   command=self.calculate,
                   style="Primary.TButton").grid(row=15, column=0, columnspan=2, pady=15)

    def create_result_table(self):
        self.project_label = ttk.Label(self.result_frame,
                                       text="",
                                       font=('Tahoma', 12, 'bold'),
                                       anchor="center")
        self.project_label.pack(fill="x", pady=5)

        columns = (
            "تاریخ شمسی",
            "تاریخ میلادی",
            "دریافت تسهیلات",
            "هزینه مدیریت",
            "هزینه تعهد",
            "سهم بهره مقدماتی",
            "مبلغ قسط",
            "سهم بهره بازپرداخت",
            "سهم اصل",
            "مانده اصل"
        )

        self.tree = ttk.Treeview(self.result_frame,
                                 columns=columns,
                                 show="headings",
                                 height=15)

        col_widths = [120, 120, 120, 90, 90, 120, 100, 120, 100, 120]
        for col, width in zip(columns, col_widths):
            self.tree.heading(col, text=col)
            self.tree.column(col, width=width, anchor="center")

        scrollbar = ttk.Scrollbar(self.result_frame,
                                  orient="vertical",
                                  command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

    def validate_inputs(self):
        try:
            return {
                'contract_amount': int(self.input_1.get()),
                'finance_rate': float(self.input_2.get() or 85),
                'interest_rate': float(self.input_3.get()),
                'insurance_rate': float(self.input_4.get() or 0),
                'perannual': int(self.input_5.get()),
                'execution_years': int(self.input_6.get()),
                'repayment_years': int(self.input_7.get()),
                'start_year': int(self.input_8.get()),
                'start_month': int(self.input_9.get()),
                'grace_years': int(self.input_10.get()),
                'grace_months': int(self.input_11.get()),
                'management_fee_rate': float(self.input_12.get() or 0),
                'commitment_fee_rate': float(self.input_13.get() or 0),
                'commitment_freq': int(self.input_14.get())
            }
        except ValueError as e:
            messagebox.showerror("خطا", f"مقادیر ورودی نامعتبر: {str(e)}")
            raise

    def calculate(self):
        try:
            inputs = self.validate_inputs()

            base_principle = inputs['contract_amount'] * (inputs['finance_rate'] / 100)
            insurance = base_principle * (inputs['insurance_rate'] / 100) * (inputs['finance_rate'] / 100)
            total_principle = base_principle + insurance

            start_date = pd.Timestamp(year=inputs['start_year'], month=inputs['start_month'], day=1)
            execution_end = start_date + relativedelta(years=inputs['execution_years'])
            grace_period = relativedelta(years=inputs['grace_years'], months=inputs['grace_months'])
            repayment_start = execution_end + grace_period

            management_fee = total_principle * (inputs['management_fee_rate'] / 100)

            drawdown_dates, drawdown_amounts = self.calculate_drawdowns(
                start_date,
                inputs['execution_years'],
                total_principle
            )

            commitment_dates, commitment_fees = self.calculate_commitments(
                start_date,
                execution_end,
                inputs['commitment_freq'],
                total_principle,
                inputs['commitment_fee_rate']
            )

            prelim_interest_dates, prelim_interests = self.calculate_preliminary_interest(
                start_date,
                execution_end,
                inputs['perannual'],
                total_principle,
                inputs['interest_rate']
            )

            repayment_dates, repayment_data = self.calculate_repayments(
                repayment_start,
                inputs['perannual'],
                inputs['repayment_years'],
                total_principle,
                inputs['interest_rate']
            )

            self.display_results(
                start_date,
                management_fee,
                drawdown_dates,
                drawdown_amounts,
                (commitment_dates, commitment_fees),
                (prelim_interest_dates, prelim_interests),
                (repayment_dates, repayment_data),
                total_principle,
                inputs['insurance_rate']
            )

        except Exception as e:
            messagebox.showerror("خطا", f"خطا در پردازش داده‌ها:\n{str(e)}")

    def convert_to_jalali(self, date):
        try:
            return jdatetime.date.fromgregorian(date=date.date()).strftime("%Y-%m")
        except:
            return ""

    def calculate_drawdowns(self, start_date, execution_years, total_principle):
        dates = []
        annual_amount = total_principle / execution_years
        amounts = [annual_amount] * execution_years

        for year in range(execution_years):
            draw_date = start_date + relativedelta(years=year)
            dates.append(draw_date)

        return dates, amounts

    def calculate_commitments(self, start_date, execution_end, freq, total_principle, rate):
        interval = 12 // freq
        dates = []
        fees = []

        total_months = (execution_end.year - start_date.year) * 12 + (execution_end.month - start_date.month)
        total_installments = total_months // interval

        installment_amount = total_principle / total_installments

        drawn_installments = 0
        current_date = start_date

        for _ in range(total_installments):
            current_date += relativedelta(months=interval)
            if current_date > execution_end:
                break

            drawn_installments += 1
            undrawn_amount = total_principle - (drawn_installments * installment_amount)

            fee = (undrawn_amount * rate / 100) / freq
            fees.append(round(fee, 2))
            dates.append(current_date)

        return dates, fees

    def calculate_preliminary_interest(self, start_date, execution_end, perannual, total_principle, rate):
        interval = 12 // perannual
        dates = []
        interests = []

        current_date = start_date + relativedelta(months=interval)
        total_months = ((execution_end.year - start_date.year) * 12 +
                        (execution_end.month - start_date.month))
        monthly_utilization = total_principle / total_months

        while current_date <= execution_end:
            months_passed = (current_date.year - start_date.year) * 12 + (current_date.month - start_date.month)
            used = monthly_utilization * months_passed
            interest = (used * rate / 100) / perannual
            dates.append(current_date)
            interests.append(round(interest, 2))
            current_date += relativedelta(months=interval)

        return dates, interests

    def calculate_repayments(self, start_date, perannual, repayment_years, total_principle, rate):
        period = repayment_years * perannual
        interval = 12 // perannual
        dates = [start_date + relativedelta(months=i * interval) for i in range(period)]

        monthly_rate = rate / 100 / perannual
        fixed_payment = (total_principle * monthly_rate * (1 + monthly_rate) ** period) / (
                (1 + monthly_rate) ** period - 1)

        remaining = total_principle
        components = []
        for _ in range(period):
            interest = remaining * monthly_rate
            principal = fixed_payment - interest
            components.append({
                'payment': round(fixed_payment, 2),
                'interest': round(interest, 2),
                'principal': round(principal, 2),
                'remaining': round(remaining - principal, 2)
            })
            remaining -= principal

        return dates, components

    def display_results(self, start_date, mgmt_fee, drawdown_dates, drawdown_amounts,
                        commit_data, prelim_data, repay_data, total_principle, insurance_rate):
        project_name = self.project_name.get() or "پروژه بدون نام"
        label_text = f"نام پروژه: {project_name}"
        if insurance_rate > 0:
            label_text += f"\nاصل وام با احتساب بیمه: {total_principle:,.0f} "

        self.project_label.config(text=label_text)
        self.tree.delete(*self.tree.get_children())

        total_drawdown = sum(drawdown_amounts)
        total_management = mgmt_fee
        total_commitment = sum(commit_data[1])
        total_prelim_interest = sum(prelim_data[1])
        total_installment = sum(item['payment'] for item in repay_data[1])
        total_repayment_interest = sum(item['interest'] for item in repay_data[1])
        total_principal = sum(item['principal'] for item in repay_data[1])

        all_events = []

        for date, amount in zip(drawdown_dates, drawdown_amounts):
            all_events.append((date, 'drawdown', amount))

        all_events.append((start_date, 'management', mgmt_fee))

        for date, fee in zip(*commit_data):
            all_events.append((date, 'commitment', fee))

        for date, interest in zip(*prelim_data):
            all_events.append((date, 'prelim', interest))

        for date, data in zip(*repay_data):
            all_events.append((date, 'repayment', data))

        all_events.sort(key=lambda x: x[0])

        remaining_principle = total_principle
        for event in all_events:
            date = event[0]
            event_type = event[1]

            jalali_date = self.convert_to_jalali(date)
            gregorian_date = date.strftime('%Y-%m') if not isinstance(date, str) else date

            values = [
                jalali_date,
                gregorian_date,
                "-", "-", "-", "-", "-", "-", "-",
                f"{remaining_principle:,.0f}"
            ]

            if event_type == 'drawdown':
                values[2] = f"{event[2]:,.0f}"
                remaining_principle -= event[2]
            elif event_type == 'management':
                values[3] = f"{event[2]:,.0f}"
            elif event_type == 'commitment':
                values[4] = f"{event[2]:,.0f}"
            elif event_type == 'prelim':
                values[5] = f"{event[2]:,.0f}"
            elif event_type == 'repayment':
                data = event[2]
                values[6] = f"{data['payment']:,.0f}"
                values[7] = f"{data['interest']:,.0f}"
                values[8] = f"{data['principal']:,.0f}"
                values[9] = f"{data['remaining']:,.0f}"
                remaining_principle = data['remaining']

            self.tree.insert("", "end", values=values)

        self.tree.insert("", "end", values=(
            "-",
            "جمع کل",
            f"{total_drawdown:,.0f}",
            f"{total_management:,.0f}",
            f"{total_commitment:,.0f}",
            f"{total_prelim_interest:,.0f}",
            f"{total_installment:,.0f}",
            f"{total_repayment_interest:,.0f}",
            f"{total_principal:,.0f}",
            "-"
        ))


if __name__ == "__main__":
    root = tk.Tk()
    app = LoanCalculatorApp(root)
    root.mainloop()