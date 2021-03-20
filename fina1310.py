import math
import xlwt


def fixed_principal(sh, n, year, int_rate):
    beg_bal = [n]
    total_pay = []
    int_paid = []
    prin_paid = []
    end_bal = []
    int_total = 0
    prin_total = 0
    total_pay_total = 0
    for i in range(year):
        int_paid.append(beg_bal[i] * int_rate)
        int_total += int_paid[i]
        prin_paid.append(n / year)
        prin_total += prin_paid[i]
        total_pay.append(int_paid[i] + prin_paid[i])
        total_pay_total += total_pay[i]
        end_bal.append(beg_bal[i] - prin_paid[i])
        if end_bal[i] != 0:
            beg_bal.append(end_bal[i])

    for i in range(year):
        int_paid.append(beg_bal[i] * int_rate)
        prin_paid.append(n / year)
        total_pay.append(int_paid[i] + prin_paid[i])
        end_bal.append(beg_bal[i] - prin_paid[i])
        if end_bal[i] != 0:
            beg_bal.append(end_bal[i])
        beg_bal[i] = round(beg_bal[i], 4)
        total_pay[i] = round(total_pay[i], 4)
        int_paid[i] = round(int_paid[i], 4)
        prin_paid[i] = round(prin_paid[i], 4)
        end_bal[i] = round(end_bal[i], 4)
        total_pay_total += total_pay[i]
        int_total += int_paid[i]
        prin_total += prin_paid[i]

    sh.write(0, 0, "Year")
    sh.write(0, 1, "Begin balance")
    sh.write(0, 2, "Total paid")
    sh.write(0, 3, "Interest paid")
    sh.write(0, 4, "Principal paid")
    sh.write(0, 5, "Ending balance")
    for i in range(year):
        sh.write(i + 1, 0, i + 1)
        sh.write(i + 1, 1, beg_bal[i])
        sh.write(i + 1, 2, total_pay[i])
        sh.write(i + 1, 3, int_paid[i])
        sh.write(i + 1, 4, prin_paid[i])
        sh.write(i + 1, 5, end_bal[i])
    sh.write(year + 2, 0, "Total")
    sh.write(year + 2, 2, total_pay_total)
    sh.write(year + 2, 3, int_total)
    sh.write(year + 2, 4, prin_total)


def fixed_payment(sh, n, year, int_rate):
    beg_bal = [n]
    total_pay = []
    int_paid = []
    prin_paid = []
    end_bal = []
    int_total = 0
    prin_total = 0
    total_pay_total = 0

    total_pay_each = n * int_rate / (1 - math.pow((1 + int_rate), year * (-1)))

    for i in range(year):
        total_pay.append(total_pay_each)
        int_paid.append(beg_bal[i] * int_rate)
        prin_paid.append(total_pay_each - int_paid[i])
        end_bal.append(beg_bal[i] - prin_paid[i])
        if end_bal[i] != 0:
            beg_bal.append(end_bal[i])
        beg_bal[i] = round(beg_bal[i], 8)
        total_pay[i] = round(total_pay[i], 8)
        int_paid[i] = round(int_paid[i], 8)
        prin_paid[i] = round(prin_paid[i], 8)
        end_bal[i] = round(end_bal[i], 8)
        total_pay_total += total_pay[i]
        int_total += int_paid[i]
        prin_total += prin_paid[i]

    sh.write(0, 0, "Year")
    sh.write(0, 1, "Begin balance")
    sh.write(0, 2, "Total paid")
    sh.write(0, 3, "Interest paid")
    sh.write(0, 4, "Principal paid")
    sh.write(0, 5, "Ending balance")
    for i in range(year):
        sh.write(i+1, 0, i+1)
        sh.write(i+1, 1, beg_bal[i])
        sh.write(i+1, 2, total_pay[i])
        sh.write(i+1, 3, int_paid[i])
        sh.write(i+1, 4, prin_paid[i])
        sh.write(i+1, 5, end_bal[i])
    sh.write(year+1, 0, "Total")
    sh.write(year+1, 2, total_pay_total)
    sh.write(year+1, 3, int_total)
    sh.write(year+1, 4, prin_total)


def amortized_loan(n, year, int_rate):
    wb = xlwt.Workbook()
    sh1 = wb.add_sheet("fixed_principal")
    fixed_principal(sh1, n, year, int_rate)
    sh2 = wb.add_sheet("fixed_payment")
    fixed_payment(sh2, n, year, int_rate)
    wb.save("amortized loan.xls")


def BV_calculate(coupon, face, int_rate, t):
    bv = 0
    bv += coupon * (1 - math.pow(1+int_rate, -t)) / int_rate
    bv += face * math.pow(1+int_rate, -t)
    return round(bv, 6)


def YTM_calculate(coupon, face, bv, t):
    max_try = 40000
    frequency = 1000
    diff_temp = 10000
    int_temp = 0
    for i in range(1, max_try * frequency+1):
        int_try = i / (frequency * 100)
        if abs(BV_calculate(coupon, face, int_try, t) - bv) < diff_temp:
            diff_temp = abs(BV_calculate(coupon, face, int_try, t) - bv)
            int_temp = int_try
        else:
            return int_temp
    return -1


if __name__ == '__main__':
    # amortized_loan(n=3900000, year=4, int_rate=0.072)
    # print(BV_calculate(coupon=39993.35, face=0, int_rate=0.10154/ 12, t=14 * 12))
    # print(YTM_calculate(coupon=15000 * 0.0742 / 2, face=15000, bv=18472, t=6 * 2), "单个时间段利率！") # face: n年后收到的大笔
    print(1.73*1.0251 / (0.10157-0.0251))
# See PyCharm help at https://www.jetbrains.com/help/pycharm/
