import re

from src.config import MajorLabels, BillLabels, AutoLabelRules


class Annotator:
    def __init__(self):
        self.labels = [label.name for label in BillLabels]
        self.rules = AutoLabelRules().rules

    def do_annotation(self, df):
        row_size = df.shape[0]
        for row in range(row_size):
            df_row = df.iloc[row]
            for label, rules in self.rules.items():
                can_annotate = self.check_rules(df_row, rules)
                if can_annotate:
                    df.iloc[row, 8] = self.bill_label_to_major_label(label).value
                    df.iloc[row, 9] = label.value
                    continue
        return df

    def check_rules(self, df, rules) -> bool:
        for condition_dict in rules:
            field_match_ok = False
            for column, condition in condition_dict.items():
                if column in ['来源', '收支', '交易对方', '商品']:
                    condition_ok = bool(re.match(condition, df[column]))
                elif column in ['金额', '修订金额']:
                    condition_ok = condition[0] < df[column] <= condition[1]
                else:
                    raise ValueError('Wrong rules format.')
                field_match_ok = field_match_ok or condition_ok
            if not field_match_ok:
                return False
        return True

    def bill_label_to_major_label(self, bill_label):
        for major_label in MajorLabels:
            if major_label.value in bill_label.value:
                return major_label
        raise ValueError(f'current label: {bill_label}, no correspond MajorLabels.')
