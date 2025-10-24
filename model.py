import pandas as pd
import numpy as np

def get_score(value, metric):
    if "Giá trị giao dịch trung bình 3 tháng" in metric:
        unit = 1000000000
        value = value/unit
        if value <1:
            return 1
        elif value < 10:
            return 2
        elif value < 50:
            return 3
        elif value < 100:
            return 4
        else:
            return 5

    if "Khối lượng giao dịch/SLCP lưu hành" in metric:
        value = value * 64
        if value < 3:
            return 1
        elif value < 10:
            return 2
        elif value < 30:
            return 3
        elif value < 65:
            return 4
        else:
            return 5


    if  metric == "Vốn hóa thị trường" :
        unit =1000000000
        value = value/unit
        if value < 500:
            return 1
        elif value < 1000:
            return 2
        elif value < 10000:
            return 3
        elif value < 50000:
            return 4
        else:
            return 5

    if ("P/e năm gần nhất" in metric) or ("P/e TTM" in metric):
        if pd.isna(value) or value <= 0: return 1
        elif value > 30: return 2
        elif value > 20: return 3
        elif value > 10: return 4
        else: return 5

    if "PE/PE ngành" in metric:
        if pd.isna(value) or value <= 0: return 1
        elif value > 1.5: return 2
        elif value > 1: return 3
        elif value > 0.5: return 4
        else: return 5

    if ("P/b năm gần nhất" in metric) or ("P/b TTM" in metric):
        if pd.isna(value) or value <= 0: return 1
        elif value > 3: return 1
        elif value > 2: return 2
        elif value > 1: return 3
        elif value > 0.5: return 4
        else: return 5


    if "PB/PB ngành" in metric:
        if pd.isna(value) or value <= 0: return 1
        elif value > 2: return 1
        elif value > 1: return 2
        elif value > 0.5: return 3
        elif value > 0.25: return 4
        else: return 5


    if "Biến động giá trong vòng 90 ngày" in metric:
        value = abs(value)*100
        if value > 100: return 1
        elif value > 75: return 2
        elif value > 50: return 3
        elif value > 25: return 4
        else: return 5

    if "ROE năm tài chính gần nhất" in metric or "ROE TTM" in metric:
        value = value *100
        if pd.isna(value) or value <= 0: return 1
        elif value <= 5: return 2
        elif value <= 10: return 3
        elif value <= 20: return 4
        else: return 5

    if "ROE cp/ROE ngành" in metric:
        if pd.isna(value) or value <= 0.5: return 1
        elif value <= 1: return 2
        elif value <= 1.5: return 3
        elif value <= 2: return 4
        else: return 5

    if "CARG 3 năm" in metric:
        if pd.isna(value) or value <= 0: return 1
        elif value <= 5: return 2
        elif value <= 10: return 3
        elif value <= 20: return 4
        else: return 5

    if "D/E năm gần nhất" in metric or "D/E TTM" in metric:
        if value < 0 or value > 2: return 1
        elif value > 1.5: return 2
        elif value > 1: return 3
        elif value > 0.5: return 4
        elif value <= 0.5: return 5
        else: return np.nan

    if "D/E cp/D/E ngành" in metric:
        if value > 2: return 1
        elif value > 1.5: return 2
        elif value > 1: return 3
        elif value > 0.5: return 4
        elif value <= 0.5: return 5
        else: return np.nan
    if "EBITDA/(nợ vay ngắn hạn+lãi vay) TTM" in metric:
        if pd.isna(value) or value < 0: return 1
        elif value <= 0.25: return 2
        elif value <= 1: return 3
        elif value <= 2: return 4
        elif value > 2: return 5
        else: return np.nan

    elif "Thanh toán hiện thời" in metric:
        if value <= 0.5: return 1
        elif value <= 1: return 2
        elif value <= 1.5: return 3
        elif value <= 2: return 4
        elif value > 2: return 5
        else: return np.nan

    elif "Thanh toán nhanh" in metric:
        if pd.isna(value) or value <= 0.25: return 1
        elif value <= 0.75: return 2
        elif value <= 1.25: return 3
        elif value <= 1.75: return 4
        else:return 5

    elif "Thanh toán tiền mặt" in metric:
        if pd.isna(value) or value < 0.05: return 1
        elif value < 0.1: return 2
        elif value < 0.25: return 3
        elif value < 0.5: return 4
        elif value >= 0.5: return 5
        else: return np.nan

    elif "Thời gian thu tiền bình quân" in metric:
        if pd.isna(value) or value < 0 or value > 360: return 1
        elif value <= 90: return 5
        elif value <= 180: return 4
        elif value <= 270: return 3
        elif value <= 360: return 2
        else: return np.nan

    elif "Thời gian hàng tồn kho" in metric:
        if pd.isna(value) or value < 0 or value > 360: return 1
        elif value <= 90: return 5
        elif value <= 180: return 4
        elif value <= 270: return 3
        elif value <= 360: return 2
        else: return np.nan

    return np.nan

def calc_cfo_score(row):
    """
    Compute CFO score based on 3 years of CFO and 3 years of LNST.
    Matches the Google Sheets IFS logic you provided.
    """

    cfo_v = row["I. Lưu chuyển tiền thuần từ hoạt động kinh doanh (GT) Hợp nhất Quý: Hàng năm Năm: 2022 Đơn vị: VND"]   # V4
    cfo_w = row["I. Lưu chuyển tiền thuần từ hoạt động kinh doanh (GT) Hợp nhất Quý: Hàng năm Năm: 2023 Đơn vị: VND"]   # W4
    cfo_x = row["I. Lưu chuyển tiền thuần từ hoạt động kinh doanh (GT) Hợp nhất Quý: Hàng năm Năm: 2024 Đơn vị: VND"]   # X4
    lnst_y = row["18. Lợi nhuận sau thuế thu nhập doanh nghiệp Hợp nhất Quý: Hàng năm Năm: 2022 Đơn vị: VND"] # Y4
    lnst_z = row["18. Lợi nhuận sau thuế thu nhập doanh nghiệp Hợp nhất Quý: Hàng năm Năm: 2023 Đơn vị: VND"] # Z4
    lnst_aa = row["18. Lợi nhuận sau thuế thu nhập doanh nghiệp Hợp nhất Quý: Hàng năm Năm: 2024 Đơn vị: VND"]# AA4

    try:
        ratio = (cfo_v + cfo_w + cfo_x) / (lnst_y + lnst_z + lnst_aa)
    except ZeroDivisionError:
        ratio = 0

    if any(pd.isna([cfo_w, cfo_x])):
        return np.nan

    if cfo_w <= 0 and cfo_x <= 0:
        return 1
    elif (cfo_w > 0) ^ (cfo_x > 0):
        return 2
    elif (cfo_w > 0) and (cfo_x > 0) and (0 < ratio <= 1):
        return 3
    elif (cfo_w > 0) and (cfo_x > 0) and (1 < ratio <= 2):
        return 4
    elif (cfo_w > 0) and (cfo_x > 0) and (2 < ratio <= 3):
        return 5
    else:
        return 5



def move_column_after(df, column_to_move, after_column):
    cols = df.columns.tolist()
    cols.remove(column_to_move)
    idx = cols.index(after_column)
    cols.insert(idx + 1, column_to_move)
    return df[cols]

def add_weighted_score(df, factors = [
    15.0, 5.0, 10.0, 4.0, 4.0, 5.0, 4.0, 4.0, 5.0, 4.0,
    3.5, 3.5, 3.0, 5.0, 3.0, 3.0, 4.0, 2.5, 2.5, 1.5,
    1.5, 2.0, 2.5, 2.5 ], exclude_cols=None):
    """
    Multiply numeric columns by given factors (percent-based),
    then create a new column 'Tổng điểm' as the sum of weighted columns.

    Parameters:
        df (pd.DataFrame): Input DataFrame.
        factors (list[float]): List of weighting factors (%).
        exclude_cols (list[str]): Columns to exclude from numeric weighting.

    Returns:
        pd.DataFrame: DataFrame with a new 'Tổng điểm' column.
    """
    import pandas as pd

    if exclude_cols is None:
        exclude_cols = ["Mã", "Tên công ty", "Sàn"]

    df_temp = df.copy()

    numeric_cols = [c for c in df_temp.select_dtypes(include="number").columns if c not in exclude_cols]

    if len(factors) != len(numeric_cols):
        raise ValueError(
            f"Number of factors ({len(factors)}) doesn't match numeric columns ({len(numeric_cols)})."
        )

    for col, factor in zip(numeric_cols, factors):
        df_temp[f"{col}_weighted"] = df_temp[col] * (factor / 100)

    df_temp["Tổng điểm"] = df_temp[[f"{col}_weighted" for col in numeric_cols]].sum(axis=1)

    df_out = df.copy()
    df_out["Tổng điểm"] = df_temp["Tổng điểm"]

    return df_out


"""
Assign letter grades (A–E) based on the 'Tổng điểm' column.
"""
def grade(score):
    if score < 3:
        return "E"
    elif  score >=3 and score < 3.25:
        return "D"
    elif score >= 3.25 and score < 3.5:
        return "C"
    elif score >= 3.5 and score < 3.75:
        return "B"
    else:
        return "A"


def add_score(df):
    df = df.copy()
    df.columns = df.columns.map(str)
    df.columns = df.columns.str.replace('\n', ' ', regex=False)
    df["Giá trị giao dịch trung bình 3 tháng"] =df[next((col for col in df.columns if "Giá trị trung bình 3 tháng" in col),None)]
    df["Khối lượng giao dịch/SLCP lưu hành"] = df["Khối lượng trung bình 3 tháng Ngày: 10/10/2025 Đơn vị: Cổ phiếu"]/df["Số cổ phiếu lưu hành Đơn vị: Cổ phiếu"]*100
    df["Vốn hóa thị trường"] =  df[next((col for col in df.columns if "Vốn hóa thị trường" in col),None)]
    df["P/e năm gần nhất"]= df[next((c for c in df.columns if "P/E" in c and "năm" in c), None)]
    df["P/e TTM"] = df[next((c for c in df.columns if "P/E" in c and "Chỉ số TTM" in c), None)]
    df["PE/PE ngành"] = df[next((c for c in df.columns if "P/E" in c and "Chỉ số TTM" in c), None)] / df["P/E ngành TTM"]
    df["P/b năm gần nhất"]= df[next((c for c in df.columns if "P/B" in c and "năm" in c), None)]
    df["P/b TTM"] =df[next((c for c in df.columns if "P/B" in c and "Chỉ số TTM" in c), None)]
    df["PB/PB ngành"] = df[next((c for c in df.columns if "P/B" in c and "Chỉ số TTM" in c), None)] / df["P/B ngành TTM"]
    df["Biến động giá trong vòng 90 ngày"]=df[next((col for col in df.columns if "% thay đổi giá 3 tháng" in col),None)]
    df["ROE năm tài chính gần nhất"] = df[next((c for c in df.columns if "ROE" in c and "năm" in c), None)]
    df["ROE TTM"]=df[next((c for c in df.columns if "ROE" in c and "Chỉ số TTM" in c), None)]
    df["ROE cp/ROE ngành"]=df[next((c for c in df.columns if "ROE" in c and "Chỉ số TTM" in c), None)] / df["ROE ngành"]
    df["CARG 3 năm"]=((df['3. Doanh thu thuần Hợp nhất Quý: Hàng năm Năm: 2024 Đơn vị: VND'] / df['3. Doanh thu thuần Hợp nhất Quý: Hàng năm Năm: 2021 Đơn vị: VND']) ** (1 / 3) - 1) * 100
    df["D/E năm gần nhất"]=df[next((c for c in df.columns if "Tổng vốn vay/ VSCH" in c and "Chỉ số năm" in c), None)]
    df["D/E TTM"]=df[next((c for c in df.columns if "Tổng vốn vay/ VSCH" in c and "Chỉ số TTM" in c), None)]
    df["D/E cp/D/E ngành"]=df["D/E TTM"]/df["D/E ngành"]
    df["EBITDA/(nợ vay ngắn hạn+lãi vay) TTM"] = df[next((c for c in df.columns if "EBITDA/ (Nợ vay ngắn hạn + lãi vay)" in c and "Chỉ số TTM" in c), None)]
    df["Thanh toán hiện thời (TTM)"] = df["Tỷ suất thanh toán hiện thời Chỉ số TTM Quý: Q2.2025 Đơn vị: Lần"]
    df["Thanh toán nhanh (TTM)"]=df["Tỷ suất thanh toán nhanh Chỉ số TTM Quý: Q2.2025 Đơn vị: Lần"]
    df["Thanh toán tiền mặt (TTM)"] = df["Tỷ suất thanh toán tiền mặt Chỉ số TTM Quý: Q2.2025 Đơn vị: Lần"]
    df["Thời gian thu tiền bình quân (TTM)"] = df["Thời gian trung bình thu tiền khách hàng Chỉ số TTM Quý: Q2.2025 Đơn vị: Ngày"]
    df["Thời gian hàng tồn kho (TTM)"] = df["Thời gian trung bình xử lý hàng tồn kho Chỉ số TTM Quý: Q2.2025 Đơn vị: Ngày"]
    df_final =pd.DataFrame()
    df_final["Mã"] = df["Mã"]
    df_final["Tên công ty"]=df["Tên công ty"]
    df_final["Sàn"] = df["Sàn"]
    for col in df.columns:
        df_final[col] = df[col].apply(lambda x: get_score(x,col))

    df_final["CFO"]=df.apply(calc_cfo_score,axis =1)
    df_final = df_final.dropna(axis=1, how='all')
    df_final["Mã"] = df["Mã"]
    df_final["Tên công ty"]=df["Tên công ty"]
    df_final["Sàn"] = df["Sàn"]
    cols = ["Mã", "Tên công ty", "Sàn"] + [c for c in df_final.columns if c not in ["Mã", "Tên công ty", "Sàn"]]
    df_final = df_final[cols]
    df_final = move_column_after(df_final, "CFO", "EBITDA/(nợ vay ngắn hạn+lãi vay) TTM")
    df_final = add_weighted_score(df_final)
    df_final["Xết loại"] = df_final["Tổng điểm"].apply(grade)

    return df_final

def to_excel(df,output):
  df_final = add_score(df)
  df_final.to_excel(output, index=False)

