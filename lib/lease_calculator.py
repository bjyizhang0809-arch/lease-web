#!/usr/bin/env python3
"""
租赁合同计算器 - 替代复杂的Excel LET公式

核心功能：
1. 计算每个合同在指定时间段的应收总额、收入总额
2. 匹配银行对账单和发票数据
3. 生成每个月的应收明细

新增功能（--aux-columns）：
4. 在Excel输出中附加辅助列（中间值和计算公式），用于排错
"""

import pandas as pd
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import calendar
import math
import sys
from pathlib import Path


class LeaseCalculator:
    """租赁合同计算器"""

    def __init__(self, contract_file: str, log_file: str = None):
        """
        初始化计算器

        Args:
            contract_file: 合同原始数据Excel文件路径
            log_file: 日志文件路径（可选）
        """
        self.contract_file = contract_file
        self.contracts_df = None
        self.bank_statements_df = None
        self.invoices_df = None
        self.max_lease_years = 6  # 默认值
        self.log_file = log_file
        self.log_lines = []  # 存储日志行
        self._load_data()
        self._detect_max_lease_years()  # 检测最大年度数

    def _load_data(self):
        """加载Excel数据"""
        try:
            # 读取合同数据
            self.contracts_df = pd.read_excel(
                self.contract_file,
                sheet_name='合同原始数据'
            )

            # 读取银行对账单
            self.bank_statements_df = pd.read_excel(
                self.contract_file,
                sheet_name='银行对账单'
            )

            # 读取发票信息
            self.invoices_df = pd.read_excel(
                self.contract_file,
                sheet_name='发票信息汇总表'
            )

            print(f"✓ 成功加载 {len(self.contracts_df)} 个合同")
            print(f"✓ 成功加载 {len(self.bank_statements_df)} 条银行对账单")
            print(f"✓ 成功加载 {len(self.invoices_df)} 条发票记录")

        except Exception as e:
            print(f"✗ 加载数据失败: {e}")
            sys.exit(1)

    def _detect_max_lease_years(self):
        """检测Excel中最大的租赁年度数"""
        import re
        max_year = 1
        for col in self.contracts_df.columns:
            match = re.match(r'保底租金第(\d+)年', col)
            if match:
                year_num = int(match.group(1))
                max_year = max(max_year, year_num)

        self.max_lease_years = max_year
        print(f"✓ 检测到最大租赁年度: {self.max_lease_years}年")

    def _log(self, message):
        """写入日志"""
        self.log_lines.append(message)
        if self.log_file:
            print(message)  # 同时打印到控制台

    def _save_log(self):
        """保存日志到文件"""
        if self.log_file:
            with open(self.log_file, 'w', encoding='utf-8') as f:
                f.write('\n'.join(self.log_lines))
            print(f"\n✓ 计算日志已保存到: {self.log_file}")

    def _get_rent_years_list(self, contract_row):
        """从合同行中提取所有年度租金列表

        Args:
            contract_row: 合同数据行

        Returns:
            list: 各年度租金列表，如[26496, 27821, ...]
        """
        rent_years_list = [contract_row['保底租金第1年（必须）']]

        for year in range(2, self.max_lease_years + 1):
            col_name = f'保底租金第{year}年'
            rent_years_list.append(contract_row.get(col_name, 0))

        return rent_years_list

    def _validate_contract_data(self, contract_row):
        """
        校验合同数据一致性：租期届满日与租金年数是否匹配

        Returns:
            list: 冲突警告字符串列表，空列表表示无冲突
        """
        warnings = []
        delivery_date = contract_row.get('交付日')
        lease_end_date = contract_row.get('租期届满日')
        customer = contract_row.get('客户名称', '')
        merchant_id = contract_row.get('商户编号', '')

        if pd.isna(delivery_date) or pd.isna(lease_end_date):
            return warnings

        delta = relativedelta(pd.to_datetime(lease_end_date) + timedelta(days=1),
                              pd.to_datetime(delivery_date))
        actual_years = delta.years + (1 if (delta.months > 0 or delta.days > 0) else 0)

        rent_years_list = self._get_rent_years_list(contract_row)
        valid_rent_years = sum(1 for r in rent_years_list if not (pd.isna(r) if not isinstance(r, str) else False) and r != 0)

        if valid_rent_years != actual_years:
            warnings.append(
                f"[数据冲突] {customer}（{merchant_id}）："
                f"租期届满日显示租期约 {actual_years} 年，"
                f"但填写了 {valid_rent_years} 年的租金数据。"
                f"计算将以租期届满日（{pd.to_datetime(lease_end_date).strftime('%Y-%m-%d')}）为准。"
            )

        return warnings

    def calculate_monthly_rent(self, delivery_date, lease_end_date, free_days,
                              rent_years_list, target_month_offset, log_detail=False,
                              return_detail=False):
        """
        计算指定月份的租金

        Args:
            delivery_date: 交付日期
            lease_end_date: 租期届满日
            free_days: 免租期天数
            rent_years_list: 各年度保底租金列表，如[26496, 27821, ...]，支持任意长度
            target_month_offset: 目标月份偏移量（0=交付当月，1=交付后第1个月...）
            log_detail: 是否记录详细日志
            return_detail: 是否返回中间值字典，用于生成辅助列

        Returns:
            return_detail=False: float 该月应收租金
            return_detail=True: (float, dict) 月租金和中间值字典
                detail dict 包含: n_free, n_eff, n_pay_total, month_days,
                is_split_year, year_num, rent_y, daily_rent,
                split_year, n1_days, n2_days, rent_y1, rent_y2, formula_str
        """
        # 构建空 detail 的辅助函数
        def _make_detail(n_free=0, n_eff=0, n_pay_total=0, month_days=0,
                         is_split_year=False, year_num=None, rent_y=None, daily_rent=None,
                         split_year=None, n1_days=None, n2_days=None,
                         rent_y1=None, rent_y2=None, formula_str=''):
            return {
                'n_free': n_free, 'n_eff': n_eff, 'n_pay_total': n_pay_total,
                'month_days': month_days, 'is_split_year': is_split_year,
                'year_num': year_num, 'rent_y': rent_y, 'daily_rent': daily_rent,
                'split_year': split_year, 'n1_days': n1_days, 'n2_days': n2_days,
                'rent_y1': rent_y1, 'rent_y2': rent_y2, 'formula_str': formula_str,
            }

        if pd.isna(delivery_date):
            if return_detail:
                return 0, _make_detail(formula_str='无交付日，租金为0')
            return 0

        # 确保日期格式正确
        if isinstance(delivery_date, str):
            delivery_date = pd.to_datetime(delivery_date)

        if isinstance(lease_end_date, str):
            lease_end_date = pd.to_datetime(lease_end_date)

        # 计算目标月份的起止日期
        target_month_start = delivery_date + relativedelta(months=target_month_offset)
        target_month_start = target_month_start.replace(day=1)
        target_month_end = target_month_start + relativedelta(months=1) - timedelta(days=1)

        if log_detail:
            self._log(f"\n    计算月份: {target_month_start.strftime('%Y-%m')} ({target_month_start.strftime('%Y-%m-%d')} 至 {target_month_end.strftime('%Y-%m-%d')})")

        # 使用实际的租赁期结束日
        lease_end = lease_end_date

        # 免租期结束日
        free_end = delivery_date + timedelta(days=int(free_days) - 1)

        # 计算目标月内的免租天数
        free_start_in_month = max(delivery_date, target_month_start)
        free_end_in_month = min(free_end, target_month_end)

        if free_start_in_month > free_end_in_month:
            n_free = 0
        else:
            n_free = (free_end_in_month - free_start_in_month).days + 1

        if log_detail and n_free > 0:
            self._log(f"      免租期: {delivery_date.strftime('%Y-%m-%d')} 至 {free_end.strftime('%Y-%m-%d')}")
            self._log(f"      本月免租天数: {n_free} 天")

        # 计算目标月内的有效天数（租赁期内）
        eff_start = max(delivery_date, target_month_start)
        eff_end = min(lease_end, target_month_end)

        if eff_start > eff_end:
            if log_detail:
                self._log(f"      不在租赁期内，租金为 0")
            if return_detail:
                return 0, _make_detail(n_free=n_free, formula_str='不在租赁期内，租金为0')
            return 0  # 不在租赁期内

        n_eff = (eff_end - eff_start).days + 1

        # 需支付租金的天数
        n_pay_total = max(0, n_eff - n_free)

        if log_detail:
            self._log(f"      有效天数（租赁期内）: {n_eff} 天")
            self._log(f"      应付租金天数: {n_pay_total} 天")

        if n_pay_total == 0:
            if log_detail:
                self._log(f"      应付天数为0，租金为 0")
            if return_detail:
                return 0, _make_detail(n_free=n_free, n_eff=n_eff, n_pay_total=0,
                                       formula_str='应付天数为0，租金为0')
            return 0

        # 使用传入的租金列表（支持任意年度数量）
        rent_years = rent_years_list

        # 各租赁年度结束日（动态生成，支持任意年度数）
        year_ends = [
            delivery_date + relativedelta(months=12*i) - timedelta(days=1)
            for i in range(1, len(rent_years) + 1)
        ]

        # 检查是否跨年度（检查除最后一年外的所有年度）
        split_year = 0
        for i, year_end in enumerate(year_ends[:-1], 1):
            if target_month_start <= year_end <= target_month_end:
                split_year = i
                break

        month_days = target_month_end.day

        if split_year == 0:
            # 不跨年度：整月属于同一租赁年
            # 计算目标月所属租赁年度
            months_diff = (target_month_start.year - delivery_date.year) * 12 + \
                         (target_month_start.month - delivery_date.month)

            if target_month_start.day < delivery_date.day:
                months_diff -= 1

            year_num = min(int(months_diff / 12) + 1, len(rent_years))
            rent_y = rent_years[year_num - 1]

            if pd.isna(rent_y) or rent_y == 0:
                if log_detail:
                    self._log(f"      租金未设置，租金为 0")
                if return_detail:
                    return 0, _make_detail(n_free=n_free, n_eff=n_eff, n_pay_total=n_pay_total,
                                           month_days=month_days, year_num=year_num,
                                           formula_str='租金未设置，租金为0')
                return 0

            daily_rent = rent_y / month_days
            monthly_rent = daily_rent * n_pay_total

            formula_str = f"{rent_y:.2f} / {month_days} × {n_pay_total} = {monthly_rent:.2f}"

            if log_detail:
                self._log(f"      租赁年度: 第{year_num}年")
                self._log(f"      年租金: {rent_y:.2f} 元")
                self._log(f"      本月天数: {month_days} 天")
                self._log(f"      日租金: {daily_rent:.2f} 元/天")
                self._log(f"      月租金 = {daily_rent:.2f} × {n_pay_total} = {monthly_rent:.2f} 元")
                self._log(f"      [公式] {formula_str}")

            if return_detail:
                return monthly_rent, _make_detail(
                    n_free=n_free, n_eff=n_eff, n_pay_total=n_pay_total,
                    month_days=month_days, is_split_year=False,
                    year_num=year_num, rent_y=rent_y, daily_rent=daily_rent,
                    formula_str=formula_str,
                )
            return monthly_rent

        else:
            # 跨年度：拆分两部分计算
            split_date = year_ends[split_year - 1]

            if log_detail:
                self._log(f"      跨年度月份，分界日: {split_date.strftime('%Y-%m-%d')}")

            # 第1部分（属于split_year年）
            n1_days = (split_date - target_month_start).days + 1
            rent_y1 = rent_years[split_year - 1]

            if pd.isna(rent_y1):
                rent_y1 = 0

            daily_rent1 = rent_y1 / month_days
            rent_part1 = daily_rent1 * n1_days

            if log_detail:
                self._log(f"      第1部分（第{split_year}年）:")
                self._log(f"        年租金: {rent_y1:.2f} 元")
                self._log(f"        天数: {n1_days} 天")
                self._log(f"        日租金: {daily_rent1:.2f} 元/天")
                self._log(f"        租金: {rent_part1:.2f} 元")

            # 第2部分（属于split_year+1年）
            n2_days = (target_month_end - split_date).days
            rent_y2 = rent_years[split_year] if split_year < len(rent_years) else 0

            if pd.isna(rent_y2):
                rent_y2 = 0

            daily_rent2 = rent_y2 / month_days
            rent_part2 = daily_rent2 * n2_days

            if log_detail:
                self._log(f"      第2部分（第{split_year + 1}年）:")
                self._log(f"        年租金: {rent_y2:.2f} 元")
                self._log(f"        天数: {n2_days} 天")
                self._log(f"        日租金: {daily_rent2:.2f} 元/天")
                self._log(f"        租金: {rent_part2:.2f} 元")

            total_rent = rent_part1 + rent_part2

            formula_str = (
                f"({rent_y1:.2f}/{month_days}×{n1_days})"
                f" + ({rent_y2:.2f}/{month_days}×{n2_days})"
                f" = {rent_part1:.2f} + {rent_part2:.2f} = {total_rent:.2f}"
            )

            if log_detail:
                self._log(f"      月租金合计 = {rent_part1:.2f} + {rent_part2:.2f} = {total_rent:.2f} 元")
                self._log(f"      [公式] {formula_str}")

            if return_detail:
                return total_rent, _make_detail(
                    n_free=n_free, n_eff=n_eff, n_pay_total=n_pay_total,
                    month_days=month_days, is_split_year=True,
                    split_year=split_year, n1_days=n1_days, n2_days=n2_days,
                    rent_y1=rent_y1, rent_y2=rent_y2,
                    formula_str=formula_str,
                )
            return total_rent

    def calculate_contract_summary(self, contract_row, start_month, end_month, log_detail=False):
        """
        计算单个合同在指定时间段的汇总数据

        Args:
            contract_row: 合同数据行
            start_month: 时间段起始月（格式：'2025-08-01'）
            end_month: 时间段结束月（格式：'2025-12-01'）
            log_detail: 是否记录详细日志

        Returns:
            dict: 包含应收总额、收入总额、银行对账单、发票对账的字典，
                  以及以 '_' 开头的辅助字段（供 process_all_contracts 提取）
        """
        customer_name = contract_row['客户名称']
        delivery_date = contract_row['交付日']
        lease_end_date = contract_row['租期届满日']

        if log_detail:
            self._log(f"\n{'='*80}")
            self._log(f"合同客户: {customer_name}")
            self._log(f"商户编号: {contract_row['商户编号']}")
            self._log(f"交付日: {delivery_date}")
            self._log(f"租期届满日: {lease_end_date}")
            self._log(f"免租期: {contract_row['免租期']} 天")

        if pd.isna(delivery_date):
            return {
                '应收总额': 0,
                '收入总额': 0,
                '银行对账单': 0,
                '发票对账': 0,
                '_合同总天数': 0,
                '_合同总应收': 0,
                '_日收入率': 0,
                '_查询期天数': 0,
                '_daily_income_rate': 0,
                '_收入计算公式': '0 / 0 × 0 = 0',
            }

        # 转换日期
        if isinstance(delivery_date, str):
            delivery_date = pd.to_datetime(delivery_date)

        if isinstance(lease_end_date, str):
            lease_end_date = pd.to_datetime(lease_end_date)

        start_date = pd.to_datetime(start_month)
        end_date = pd.to_datetime(end_month)
        end_date = end_date + relativedelta(months=1) - timedelta(days=1)  # 月末

        if log_detail:
            self._log(f"查询时间段: {start_date.strftime('%Y-%m-%d')} 至 {end_date.strftime('%Y-%m-%d')}")

        # 计算合同总月数：从交付日所在月到租期届满日所在月（包括首尾）
        total_months = (lease_end_date.year - delivery_date.year) * 12 + \
                      (lease_end_date.month - delivery_date.month) + 1

        if log_detail:
            self._log(f"\n【应收总额计算】")
            self._log(f"  公式: 月租金 = 年租金 / 月天数 × 应付天数")
            self._log(f"  跨年公式: (年1租金/月天数×年1天数) + (年2租金/月天数×年2天数)")
            self._log(f"  按月逐一计算租金，考虑免租期和跨年度情况")

        # 计算时间段内每个月的应收
        total_receivable = 0
        current_date = start_date.replace(day=1)

        # 提取租金列表（支持动态年度数）
        rent_years_list = self._get_rent_years_list(contract_row)

        if log_detail:
            for i, rent in enumerate(rent_years_list, 1):
                self._log(f"  第{i}年租金: {rent:.2f} 元")

        while current_date <= end_date:
            # 计算从交付日到当前月的月份偏移
            months_offset = (current_date.year - delivery_date.year) * 12 + \
                           (current_date.month - delivery_date.month)

            monthly_rent = self.calculate_monthly_rent(
                delivery_date,
                lease_end_date,
                contract_row['免租期'],
                rent_years_list,
                months_offset,
                log_detail=log_detail
            )

            total_receivable += monthly_rent
            current_date += relativedelta(months=1)

        if log_detail:
            self._log(f"\n  应收总额合计: {total_receivable:.2f} 元")

        # 计算收入总额（按新逻辑）
        if log_detail:
            self._log(f"\n【收入总额计算】")
            self._log(f"  公式: 收入总额 = 合同总应收 / 合同总天数 × 查询期天数")
            self._log(f"  基于日收入率，平滑分摊到合同期内每一天")

        # 1. 计算合同内总时间的应收总额
        total_contract_receivable = 0
        for month_offset in range(total_months):
            monthly_rent = self.calculate_monthly_rent(
                delivery_date,
                lease_end_date,
                contract_row['免租期'],
                rent_years_list,
                month_offset,
                log_detail=False  # 不记录详细日志
            )
            total_contract_receivable += monthly_rent

        # 2. 计算合同内的总天数（使用实际租期届满日）
        lease_end = lease_end_date
        total_contract_days = (lease_end - delivery_date).days + 1

        # 3. 计算收入日租金
        if total_contract_days > 0:
            daily_income_rate = total_contract_receivable / total_contract_days
        else:
            daily_income_rate = 0

        if log_detail:
            self._log(f"  合同总应收: {total_contract_receivable:.2f} 元")
            self._log(f"  合同总天数: {total_contract_days} 天")
            self._log(f"  日收入率: {daily_income_rate:.4f} 元/天")

        # 4. 计算这段时间在合同期内的天数
        # 时间段的实际起止日期与合同期的交集
        period_start = max(start_date, delivery_date)
        period_end = min(end_date, lease_end)

        if period_start <= period_end:
            days_in_period = (period_end - period_start).days + 1
        else:
            days_in_period = 0

        # 5. 计算一定时间段内的收入总额
        total_income = daily_income_rate * days_in_period

        if log_detail:
            self._log(f"  查询期内天数: {days_in_period} 天")
            self._log(f"  收入总额 = {total_contract_receivable:.2f} / {total_contract_days} × {days_in_period} = {total_income:.2f} 元")

        # 匹配银行对账单
        bank_total = self._match_bank_statements(customer_name, start_date, end_date)

        # 匹配发票
        invoice_total = self._match_invoices(customer_name, start_date, end_date)

        if log_detail:
            self._log(f"\n【匹配结果】")
            self._log(f"  银行对账单: {bank_total:.2f} 元")
            self._log(f"  发票对账: {invoice_total:.2f} 元")

        income_formula = (
            f"{total_contract_receivable:.2f} / {total_contract_days}"
            f" × {days_in_period} = {total_income:.2f}"
        )

        return {
            '应收总额': round(total_receivable, 2),
            '收入总额': round(total_income, 2),
            '银行对账单': round(bank_total, 2),
            '发票对账': round(invoice_total, 2),
            # 辅助字段（以 _ 开头，供 process_all_contracts 按需提取）
            '_合同总天数': total_contract_days,
            '_合同总应收': round(total_contract_receivable, 2),
            '_日收入率': round(daily_income_rate, 4),
            '_查询期天数': days_in_period,
            '_daily_income_rate': daily_income_rate,
            '_收入计算公式': income_formula,
        }

    def _match_bank_statements(self, customer_name, start_date, end_date):
        """匹配银行对账单"""
        try:
            # 筛选符合条件的记录
            mask = (
                (self.bank_statements_df['对方户名'] == customer_name) &
                (pd.to_datetime(self.bank_statements_df['交易时间']) >= start_date) &
                (pd.to_datetime(self.bank_statements_df['交易时间']) <= end_date)
            )

            matched = self.bank_statements_df[mask]
            return matched['贷方发生额（收入）'].sum()
        except Exception as e:
            print(f"警告：匹配银行对账单失败 - {e}")
            return 0

    def _match_invoices(self, customer_name, start_date, end_date):
        """匹配发票"""
        try:
            # 筛选符合条件的记录
            mask = (
                (self.invoices_df['购买方名称'] == customer_name) &
                (pd.to_datetime(self.invoices_df['开票日期']) >= start_date) &
                (pd.to_datetime(self.invoices_df['开票日期']) <= end_date)
            )

            matched = self.invoices_df[mask]
            return matched['价税合计'].sum()
        except Exception as e:
            print(f"警告：匹配发票失败 - {e}")
            return 0

    def calculate_monthly_breakdown(self, contract_row, start_month, end_month,
                                    with_aux=False):
        """
        计算单个合同在时间段内每个月的应收明细

        Args:
            with_aux: 是否附加辅助列（中间值和计算公式），用于排错

        Returns:
            list: 每个月的应收金额列表，with_aux=True 时附加辅助字段
        """
        delivery_date = contract_row['交付日']
        lease_end_date = contract_row['租期届满日']

        if pd.isna(delivery_date):
            return []

        if isinstance(lease_end_date, str):
            lease_end_date = pd.to_datetime(lease_end_date)

        start_date = pd.to_datetime(start_month)
        end_date = pd.to_datetime(end_month)

        # 提取租金列表（支持动态年度数）
        rent_years_list = self._get_rent_years_list(contract_row)

        monthly_list = []
        current_date = start_date.replace(day=1)

        while current_date <= end_date:
            months_offset = (current_date.year - delivery_date.year) * 12 + \
                           (current_date.month - delivery_date.month)

            if with_aux:
                monthly_rent, detail = self.calculate_monthly_rent(
                    delivery_date,
                    lease_end_date,
                    contract_row['免租期'],
                    rent_years_list,
                    months_offset,
                    return_detail=True
                )

                # 处理跨年度时的年租金和日租金显示
                if detail['is_split_year']:
                    year_num_str = f"第{detail['split_year']}/{detail['split_year']+1}年"
                    md = detail['month_days'] if detail['month_days'] else 1
                    rent_y_str = f"{detail['rent_y1']:.2f}/{detail['rent_y2']:.2f}"
                    daily_rent_str = f"{detail['rent_y1']/md:.2f}/{detail['rent_y2']/md:.2f}"
                else:
                    year_num_str = f"第{detail['year_num']}年" if detail['year_num'] else '-'
                    rent_y_str = f"{detail['rent_y']:.2f}" if detail['rent_y'] is not None else '-'
                    daily_rent_str = f"{detail['daily_rent']:.2f}" if detail['daily_rent'] is not None else '-'

                monthly_list.append({
                    '月份': current_date.strftime('%Y-%m'),
                    '应收金额': round(monthly_rent, 2),
                    '免租天数': detail['n_free'],
                    '有效天数': detail['n_eff'],
                    '应付天数': detail['n_pay_total'],
                    '月天数': detail['month_days'],
                    '租赁年度': year_num_str,
                    '年租金': rent_y_str,
                    '日租金': daily_rent_str,
                    '是否跨年度': '是' if detail['is_split_year'] else '否',
                    '计算公式': detail['formula_str'],
                })
            else:
                monthly_rent = self.calculate_monthly_rent(
                    delivery_date,
                    lease_end_date,
                    contract_row['免租期'],
                    rent_years_list,
                    months_offset
                )

                monthly_list.append({
                    '月份': current_date.strftime('%Y-%m'),
                    '应收金额': round(monthly_rent, 2)
                })

            current_date += relativedelta(months=1)

        return monthly_list

    def calculate_monthly_income_breakdown(self, contract_row, start_month, end_month,
                                           daily_income_rate, with_aux=False):
        """
        计算单个合同在时间段内每个月的收入明细

        Args:
            contract_row: 合同数据行
            start_month: 时间段起始月
            end_month: 时间段结束月
            daily_income_rate: 日收入率
            with_aux: 是否附加辅助列（日收入率、本月合同天数、计算公式）

        Returns:
            list: 每个月的收入金额列表
        """
        delivery_date = contract_row['交付日']
        lease_end_date = contract_row['租期届满日']

        if pd.isna(delivery_date):
            return []

        # 确保日期格式正确
        if isinstance(delivery_date, str):
            delivery_date = pd.to_datetime(delivery_date)
        if isinstance(lease_end_date, str):
            lease_end_date = pd.to_datetime(lease_end_date)

        start_date = pd.to_datetime(start_month)
        end_date = pd.to_datetime(end_month)
        end_date = end_date + relativedelta(months=1) - timedelta(days=1)  # 月末

        monthly_list = []
        current_date = start_date.replace(day=1)

        while current_date <= pd.to_datetime(end_month):
            # 计算当前月的起止日期
            month_start = current_date
            month_end = current_date + relativedelta(months=1) - timedelta(days=1)

            # 计算该月在合同期内的天数
            period_start = max(month_start, delivery_date)
            period_end = min(month_end, lease_end_date)

            if period_start <= period_end:
                days_in_month = (period_end - period_start).days + 1
                monthly_income = daily_income_rate * days_in_month
            else:
                days_in_month = 0
                monthly_income = 0

            item = {
                '月份': current_date.strftime('%Y-%m'),
                '收入金额': round(monthly_income, 2),
            }

            if with_aux:
                item['日收入率'] = round(daily_income_rate, 4)
                item['本月合同天数'] = days_in_month
                item['计算公式'] = (
                    f"{daily_income_rate:.4f} × {days_in_month} = {monthly_income:.2f}"
                )

            monthly_list.append(item)
            current_date += relativedelta(months=1)

        return monthly_list

    def process_all_contracts(self, start_month, end_month, output_dir='.',
                              enable_log=False, aux_columns=False):
        """
        处理所有合同，生成三个独立的Excel报表

        Args:
            start_month: 时间段起始月（格式：'2025-08-01'）
            end_month: 时间段结束月（格式：'2025-12-01'）
            output_dir: 输出目录路径（默认为当前目录）
            enable_log: 是否启用详细日志
            aux_columns: 是否在Excel输出中添加辅助列（中间值和计算公式），用于排错

        Returns:
            tuple: (汇总DataFrame, 应收明细DataFrame, 收入明细DataFrame)
        """
        print(f"\n开始计算时间段: {start_month} 至 {end_month}")
        if aux_columns:
            print("  模式: 启用辅助列（排错模式）")
        print("="*60)

        if enable_log:
            self._log(f"租赁合同计算日志")
            self._log(f"计算时间段: {start_month} 至 {end_month}")
            self._log(f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

        summary_results = []
        monthly_receivables_results = []
        monthly_income_results = []

        for idx, row in self.contracts_df.iterrows():
            customer_name = row['客户名称']
            merchant_id = row['商户编号']

            print(f"\n处理合同 {idx+1}/{len(self.contracts_df)}: {customer_name} ({merchant_id})")

            # 校验合同数据一致性
            validation_warnings = self._validate_contract_data(row)
            for w in validation_warnings:
                print(f"  ⚠️  {w}")
                self._log(w)

            # 计算汇总数据（返回值包含辅助字段 _daily_income_rate 等）
            summary = self.calculate_contract_summary(row, start_month, end_month, log_detail=enable_log)

            # 从 summary 中取出日收入率（避免重复计算）
            daily_income_rate = summary['_daily_income_rate']

            # 计算应收月度明细
            monthly_breakdown = self.calculate_monthly_breakdown(
                row, start_month, end_month, with_aux=aux_columns
            )

            # 计算收入月度明细
            monthly_income_breakdown = self.calculate_monthly_income_breakdown(
                row, start_month, end_month, daily_income_rate, with_aux=aux_columns
            )

            # 汇总数据（用于lease.xlsx）
            lease_end_date = row['租期届满日']
            summary_result = {
                '客户名称': customer_name,
                '商户编号': merchant_id,
                '交付日': row['交付日'],
                '租期届满日': lease_end_date,  # 始终输出
                '免租期': row['免租期'],
                '应收总额': summary['应收总额'],
                '收入总额': summary['收入总额'],
                '银行对账单': summary['银行对账单'],
                '发票对账': summary['发票对账'],
                '数据备注': ' | '.join(validation_warnings) if validation_warnings else ''
            }

            # 辅助列（仅 --aux-columns 时添加）
            if aux_columns:
                summary_result['合同总天数'] = summary['_合同总天数']
                summary_result['合同总应收'] = summary['_合同总应收']
                summary_result['日收入率'] = summary['_日收入率']
                summary_result['查询期天数'] = summary['_查询期天数']
                summary_result['收入计算公式'] = summary['_收入计算公式']

            summary_results.append(summary_result)

            # 应收月度明细数据（用于single.xlsx）
            for month_data in monthly_breakdown:
                monthly_result = {
                    '客户名称': customer_name,
                    '商户编号': merchant_id,
                    '月份': month_data['月份'],
                    '应收金额': month_data['应收金额'],
                }
                if aux_columns:
                    monthly_result['免租天数'] = month_data.get('免租天数', '')
                    monthly_result['有效天数'] = month_data.get('有效天数', '')
                    monthly_result['应付天数'] = month_data.get('应付天数', '')
                    monthly_result['月天数'] = month_data.get('月天数', '')
                    monthly_result['租赁年度'] = month_data.get('租赁年度', '')
                    monthly_result['年租金'] = month_data.get('年租金', '')
                    monthly_result['日租金'] = month_data.get('日租金', '')
                    monthly_result['是否跨年度'] = month_data.get('是否跨年度', '')
                    monthly_result['计算公式'] = month_data.get('计算公式', '')
                monthly_receivables_results.append(monthly_result)

            # 收入月度明细数据（用于income.xlsx）
            for month_data in monthly_income_breakdown:
                monthly_income_result = {
                    '客户名称': customer_name,
                    '商户编号': merchant_id,
                    '月份': month_data['月份'],
                    '收入金额': month_data['收入金额'],
                }
                if aux_columns:
                    monthly_income_result['日收入率'] = month_data.get('日收入率', '')
                    monthly_income_result['本月合同天数'] = month_data.get('本月合同天数', '')
                    monthly_income_result['计算公式'] = month_data.get('计算公式', '')
                monthly_income_results.append(monthly_income_result)

            print(f"  应收总额: {summary['应收总额']:.2f}")
            print(f"  收入总额: {summary['收入总额']:.2f}")
            print(f"  银行对账单: {summary['银行对账单']:.2f}")
            print(f"  发票对账: {summary['发票对账']:.2f}")

        # 转换为DataFrame
        summary_df = pd.DataFrame(summary_results)
        monthly_receivables_df = pd.DataFrame(monthly_receivables_results)
        monthly_income_df = pd.DataFrame(monthly_income_results)

        # 保存到三个独立的Excel文件
        ts = datetime.now().strftime('%Y%m%d%H%M%S')
        lease_file = Path(output_dir) / f'{ts}-lease.xlsx'
        single_file = Path(output_dir) / f'{ts}-single.xlsx'
        income_file = Path(output_dir) / f'{ts}-income.xlsx'

        summary_df.to_excel(lease_file, index=False)
        print(f"\n✓ 汇总数据已保存到: {lease_file}")

        monthly_receivables_df.to_excel(single_file, index=False)
        print(f"✓ 应收月度明细已保存到: {single_file}")

        monthly_income_df.to_excel(income_file, index=False)
        print(f"✓ 收入月度明细已保存到: {income_file}")

        # 保存日志
        if enable_log:
            self._save_log()

        print("\n" + "="*60)
        print("计算完成！")

        return summary_df, monthly_receivables_df, monthly_income_df


def main():
    """主函数"""
    import argparse

    parser = argparse.ArgumentParser(description='租赁合同计算器')
    parser.add_argument('contract_file', help='合同原始数据Excel文件路径')
    parser.add_argument('--start', required=True, help='时间段起始月（格式：2025-08-01）')
    parser.add_argument('--end', required=True, help='时间段结束月（格式：2025-12-01）')
    parser.add_argument('--output-dir', default='.', help='输出目录路径（默认为当前目录）')
    parser.add_argument('--log', help='日志文件路径（可选，如：log.txt）')
    parser.add_argument('--aux-columns', action='store_true',
                        help='在Excel输出中添加辅助列（中间值和计算公式），用于排错')

    args = parser.parse_args()

    # 创建计算器
    calculator = LeaseCalculator(args.contract_file, log_file=args.log)

    # 处理所有合同，生成三个Excel文件
    summary_df, monthly_receivables_df, monthly_income_df = calculator.process_all_contracts(
        args.start,
        args.end,
        args.output_dir,
        enable_log=(args.log is not None),
        aux_columns=args.aux_columns
    )

    # 显示汇总统计
    print("\n汇总统计:")
    print(f"  总应收: {summary_df['应收总额'].sum():.2f}")
    print(f"  总收入: {summary_df['收入总额'].sum():.2f}")
    print(f"  总银行对账单: {summary_df['银行对账单'].sum():.2f}")
    print(f"  总发票对账: {summary_df['发票对账'].sum():.2f}")
    print(f"\n应收月度明细记录数: {len(monthly_receivables_df)}")
    print(f"收入月度明细记录数: {len(monthly_income_df)}")


if __name__ == '__main__':
    main()
