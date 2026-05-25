from pathlib import Path
import csv
import xlsxwriter


ROOT = Path(__file__).resolve().parents[1]
XLSX_PATH = ROOT / "text_replace_inventory.xlsx"
CSV_PATH = ROOT / "text_replace_inventory.csv"

headers = ["구분", "세부영역", "키", "현재문구", "교체문구", "비고"]

rows = [
    ["공통", "문서/앱 제목", "app_title", "우리의 1주년", "", "index title, apple-mobile-web-app-title"],
    ["상단바", "브랜드 문구", "brandTitle", "2025.05.27 하린이와 홍래 사랑 진행 중 ♥", "", "상단 고정바"],
    ["내비게이션", "요약 링크", "nav_summary", "요약", "", ""],
    ["내비게이션", "앨범 링크", "nav_album", "앨범", "", ""],
    ["내비게이션", "타임라인 링크", "nav_timeline", "타임라인", "", ""],
    ["내비게이션", "편지 링크", "nav_letter", "편지", "", ""],
    ["메인", "키커", "hero_kicker", "♥ 뭉리니 스크랩북", "", ""],
    ["메인", "큰 제목", "heroTitle", "우리의 아름다운 1년", "", ""],
    ["메인", "서브 문구 1", "heroSubtitle", "함께한 계절들이 모여, 오늘의 우리가 됐어.", "", ""],
    ["메인", "서브 문구 2", "hero_subcopy", "사진과 메모를 넘기듯, 지난 1년을 천천히 꺼내보는 작은 기록장.", "", ""],
    ["메인", "버튼", "hero_button_album", "앨범 열기", "", ""],
    ["메인", "버튼", "hero_button_letter", "편지 읽기", "", ""],
    ["메인 5장 카드", "배지", "hero_stack_badge", "뭉리니 순간", "", "CSS content"],
    ["메인 5장 카드", "카드 라벨 1", "hero_card_1", "분홍빛 밤", "", ""],
    ["메인 5장 카드", "카드 라벨 2", "hero_card_2", "파란 장면", "", ""],
    ["메인 5장 카드", "카드 라벨 3", "hero_card_3", "카페의 우리", "", ""],
    ["메인 5장 카드", "카드 라벨 4", "hero_card_4", "햇살 나들이", "", ""],
    ["메인 5장 카드", "카드 라벨 5", "hero_card_5", "첫 화면", "", ""],
    ["요약", "타이틀", "summary_title", "Memory Summary", "", ""],
    ["요약", "설명", "summary_desc", "처음의 컨셉처럼, 우리가 함께 쌓아온 시간을 네 장의 카드로만 담백하게 모았어.", "", ""],
    ["요약", "카드 라벨", "summary_days_label", "함께한 날", "", ""],
    ["요약", "카드 값", "summary_days_value", "D+{days}", "", "동적 값"],
    ["요약", "카드 하단", "summary_days_stamp", "2025.05.27 - Today", "", ""],
    ["요약", "카드 라벨", "summary_months_label", "계절처럼 쌓인 달", "", ""],
    ["요약", "카드 값", "summary_months_value", "12", "", ""],
    ["요약", "카드 라벨", "summary_moments_label", "다시 보고 싶은 순간", "", ""],
    ["요약", "카드 값", "summary_moments_value", "∞", "", ""],
    ["요약", "카드 라벨", "summary_anniversary_label", "첫 번째 기념일", "", ""],
    ["요약", "카드 값", "summary_anniversary_value", "1st", "", ""],
    ["앨범", "타이틀", "album_title", "Photo Album", "", ""],
    ["앨범", "설명", "album_desc", "계절별로 묶은 우리의 앨범. 폴라로이드처럼 눌러서 더 크게 볼 수 있어.", "", ""],
]

seasons = [
    (
        "여름",
        "summer",
        "2025 여름",
        "여름, 우리라는 계절의 시작",
        "첫 계절",
        "5월 27일 이후의 첫 계절. 서툴고 뜨거웠지만, 그래서 더 선명했던 시작.",
        [
            "처음이라 더 선명했던 우리. 서툰 마음까지 여름처럼 반짝였던 날.",
            "익숙해지는 속도보다 좋아지는 마음이 조금 더 빨랐던 계절.",
            "더운 공기 속에서도 같이 웃으면 이상하게 시원했던 날.",
            "처음 맞는 여름이라 모든 장면이 조금 더 크게 남았어.",
            "우리라는 이름이 자연스러워지기 시작한 계절.",
        ],
    ),
    (
        "가을",
        "autumn",
        "2025 가을",
        "가을, 조금 더 우리답게",
        "제주도와 테니스",
        "선선한 제주 바람 속에 잊지 못할 추억을 많이 남겼고, 함께 시작한 테니스까지 유난히 좋았던 계절.",
        [
            "제주 바람 속에서 남긴 장면들. 움직이는 순간까지 오래 보고 싶은 가을.",
            "선선한 날씨와 테니스, 그리고 우리답게 웃던 시간이 가득했던 계절.",
            "조금 느린 바람과 조금 더 가까워진 마음이 닮아 있던 날.",
            "같이 시작한 취미가 우리 사이에 또 하나의 이야기가 됐어.",
            "가을빛 아래에서 우리답게 걷고 웃고 기억했던 순간.",
        ],
    ),
    (
        "겨울",
        "winter",
        "2025 겨울",
        "겨울, 따뜻한 쪽으로",
        "차, 캠핑, 스키장",
        "함께 여행할 차가 생기고, 캠핑이라는 취미도 만들고, 스키장까지 다녀오며 추억을 차곡차곡 쌓았던 계절.",
        [
            "차가 생기고 우리의 이동 반경도, 추억의 크기도 조금 더 넓어졌어.",
            "추운 계절인데도 함께 있으면 묘하게 따뜻했던 순간.",
            "캠핑이라는 취미가 생기고, 평범한 밤도 특별한 장면이 됐어.",
            "작은 케이크 하나에도 우리의 표정이 오래 남았던 날.",
            "겨울 끝자락에 남긴 다정한 기념일, 그래서 더 소중한 사진.",
        ],
    ),
    (
        "봄",
        "spring",
        "2026 봄",
        "봄, 다시 피어난 우리",
        "차박, 캠핑, 나들이",
        "차박, 캠핑, 나들이로 이곳저곳 많이 다녔고, 업그레이드된 캠핑 장비 덕분에 앞으로가 더 기대되는 계절.",
        [
            "다시 따뜻해진 날씨처럼 우리도 더 자주 밖으로 나섰던 봄.",
            "나들이와 차박, 캠핑으로 주말마다 작은 여행을 만들던 시간.",
            "햇살이 좋은 날엔 사진도 마음도 조금 더 밝아졌어.",
            "업그레이드된 캠핑 장비만큼 앞으로의 계절도 기대되던 순간.",
            "풍경보다 먼저 떠오르는 건 결국 옆에 있던 너의 표정.",
        ],
    ),
]

for season, sid, date, title, place, memo, captions in seasons:
    rows.extend(
        [
            ["계절별 앨범", f"{season} 날짜", f"{sid}_date", date, "", ""],
            ["계절별 앨범", f"{season} 제목", f"{sid}_title", title, "", ""],
            ["계절별 앨범", f"{season} 장소/카테고리", f"{sid}_place", place, "", ""],
            ["계절별 앨범", f"{season} 요약문구", f"{sid}_memo", memo, "", "앨범 카드/팝업 요약"],
        ]
    )
    for index, caption in enumerate(captions, start=1):
        exists = season in ["겨울", "봄"] or index <= 2
        note = "현재 코드에 있음" if exists else "샘플 텍스트, 사진 추가 시 사용"
        rows.append(["계절별 앨범 사진문구", f"{season} {index}번 사진 글귀", f"{sid}_photo_{index}_caption", caption, "", note])

rows.extend(
    [
        ["현재 영역", "타이틀", "current_title", "그리고, 현재의 우리", "", ""],
        ["현재 영역", "업로드 상태", "current_status_local", "사진을 이 기기에 담아뒀어.", "", "Supabase 미설정 fallback"],
        ["현재 영역", "업로드 상태", "current_status_start", "업로드 중...", "", ""],
        ["현재 영역", "업로드 상태", "current_status_progress", "업로드 중 {progress}%", "", "동적 값"],
        ["현재 영역", "업로드 상태", "current_status_done", "업로드 완료", "", ""],
        ["현재 영역", "업로드 오류", "current_status_size_error", "12MB 이하 사진으로 다시 선택해줘.", "", ""],
        ["현재 영역", "업로드 오류", "current_status_fail", "업로드 실패 ({detail})", "", "동적 값"],
        ["타임라인", "타이틀", "timeline_title", "Timeline", "", ""],
        ["타임라인", "설명", "timeline_desc", "시간 순서대로 꺼내보는 우리의 페이지. 작은 하트마다 하나의 장면이 있어.", "", ""],
        ["타임라인", "항목 1 날짜", "timeline_1_date", "2025.05.27", "", ""],
        ["타임라인", "항목 1 제목", "timeline_1_title", "우리가 시작된 날", "", ""],
        ["타임라인", "항목 1 설명", "timeline_1_desc", "조금 서툴렀지만 오래 기억에 남은 시작.", "", ""],
        ["타임라인", "항목 2 날짜", "timeline_2_date", "2025 여름", "", ""],
        ["타임라인", "항목 2 제목", "timeline_2_title", "우리라는 계절의 시작", "", ""],
        ["타임라인", "항목 2 설명", "timeline_2_desc", "처음이라 더 선명했던 날들이 쌓였어.", "", ""],
        ["타임라인", "항목 3 날짜", "timeline_3_date", "2025 가을", "", ""],
        ["타임라인", "항목 3 제목", "timeline_3_title", "제주와 테니스", "", ""],
        ["타임라인", "항목 3 설명", "timeline_3_desc", "선선한 날씨와 함께 우리다운 추억이 많아졌어.", "", ""],
        ["타임라인", "항목 4 날짜", "timeline_4_date", "2025 겨울", "", ""],
        ["타임라인", "항목 4 제목", "timeline_4_title", "차와 캠핑, 스키장", "", ""],
        ["타임라인", "항목 4 설명", "timeline_4_desc", "함께 떠날 수 있는 길이 더 넓어진 계절.", "", ""],
        ["타임라인", "항목 5 날짜", "timeline_5_date", "2026 봄", "", ""],
        ["타임라인", "항목 5 제목", "timeline_5_title", "다시 피어난 우리", "", ""],
        ["타임라인", "항목 5 설명", "timeline_5_desc", "차박과 캠핑, 나들이로 앞으로가 더 기대되는 봄.", "", ""],
        ["편지", "제목", "letter_title", "To. 사랑하는 너에게", "", ""],
        ["편지", "본문 1", "letter_body_1", "우리 벌써 1년이라니, 시간은 빠른데 나는 아직도 네가 웃으면 잠깐 멈칫하는 사람으로 살고 있어.", "", ""],
        ["편지", "본문 2", "letter_body_2", "지난 1년 동안 같이 웃고, 먹고, 걷고, 가끔은 아무 말 없이 있어도 편했던 순간들이 생각보다 많이 쌓였더라.", "", ""],
        ["편지", "본문 3", "letter_body_3", "앞으로도 우리답게, 가볍게 웃고 오래 함께하자.", "", ""],
        ["편지", "서명", "letter_signature", "From. 홍래", "", ""],
        ["마지막", "제목", "final_title", "앞으로의 우리에게", "", ""],
        ["마지막", "문구", "final_text", "아직 채워지지 않은 페이지가 더 많아서 좋아.", "", ""],
        ["마지막", "버튼", "final_button", "다시 처음으로", "", ""],
        ["모달", "닫기 버튼", "modal_close", "닫기", "", ""],
        ["푸터", "문구", "footer_text", "Made for our first anniversary", "", ""],
    ]
)


def write_xlsx() -> None:
    workbook = xlsxwriter.Workbook(XLSX_PATH)
    worksheet = workbook.add_worksheet("REPLACE")
    header_format = workbook.add_format(
        {"bold": True, "bg_color": "#F4A7A3", "font_color": "#3D2C2E", "border": 1, "align": "center", "valign": "vcenter"}
    )
    cell_format = workbook.add_format({"border": 1, "valign": "top", "text_wrap": True})
    key_format = workbook.add_format({"border": 1, "valign": "top", "text_wrap": True, "font_color": "#7B4E45"})

    for col, header in enumerate(headers):
        worksheet.write(0, col, header, header_format)
    for row_index, row in enumerate(rows, start=1):
        for col_index, value in enumerate(row):
            worksheet.write(row_index, col_index, value, key_format if col_index == 2 else cell_format)

    worksheet.freeze_panes(1, 0)
    worksheet.autofilter(0, 0, len(rows), len(headers) - 1)
    for col, width in enumerate([18, 24, 28, 62, 62, 34]):
        worksheet.set_column(col, col, width)
    workbook.close()


def write_csv() -> None:
    with CSV_PATH.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.writer(file)
        writer.writerow(headers)
        writer.writerows(rows)


if __name__ == "__main__":
    write_xlsx()
    write_csv()
    print(XLSX_PATH)
    print(CSV_PATH)
    print(len(rows))
