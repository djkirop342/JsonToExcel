import json
import pandas as pd
from tqdm import tqdm

# JSON 파일 경로
json_file_path = "data.json"  # JSON 파일 경로
excel_file_path = "output.xlsx"  # 저장할 Excel 파일 경로

# JSON 파일 읽기
with open(json_file_path, "r", encoding="utf-8") as file:
    json_data = json.load(file)

# 데이터 개수 확인
total_records = len(json_data)
print(f"총 데이터 개수: {format(total_records, ',')}개")

# JSON 데이터를 DataFrame으로 변환
df = pd.DataFrame(json_data)

# Excel 저장 중 진행 상태 표시
with pd.ExcelWriter(excel_file_path, engine="xlsxwriter") as writer:
    # tqdm 프로그래스바 설정
    with tqdm(
        total=total_records,
        desc="진행상태",
        unit="row",
        bar_format="{l_bar}{bar}| {n_fmt}/{total_fmt} [{elapsed}/{remaining}]"
    ) as pbar:
        # 1000행 단위로 저장
        for start_row in range(0, total_records, 1000): 
            end_row = min(start_row + 1000, total_records)
            # 첫 번째 배치에는 헤더를 포함하고, 이후 배치에서는 헤더 제외
            df.iloc[start_row:end_row].to_excel(
                writer,
                index=False,
                header=(start_row == 0),  # 첫 번째 배치만 헤더 포함
                startrow=start_row if start_row > 0 else 0,  # 항상 숫자를 전달
                sheet_name="sheet 1"
            )
            pbar.update(end_row - start_row)

print(f"Excel 파일이 생성되었습니다: {excel_file_path}")
