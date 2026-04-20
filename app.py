import os, json, math
from flask import Flask, jsonify, request
from flask_cors import CORS
import gspread
from google.oauth2.service_account import Credentials

app = Flask(__name__)
CORS(app)

# ── 구글 시트 연결 ──────────────────────────────────────────
SHEET_ID = '15AnatVs4sauXt2FXLkqVzpmyTtVdpGRNTjmVx7VKoZ0'
MAIN_SHEET = '펀드 투자검토보고서 업데이트 현황_20260203'
HEADER_ROW = 10  # 헤더가 10번째 행

SCOPES = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']

_gc = None
_sh = None

def get_sheet():
    global _gc, _sh
    if _gc is None:
        # Render 환경변수에서 JSON 키 읽기
        creds_json = os.environ.get('GOOGLE_CREDENTIALS')
        if creds_json:
            creds_data = json.loads(creds_json)
        else:
            # 로컬 테스트용
            with open('credentials.json', 'r') as f:
                creds_data = json.load(f)
        creds = Credentials.from_service_account_info(creds_data, scopes=SCOPES)
        _gc = gspread.authorize(creds)
    if _sh is None:
        _sh = _gc.open_by_key(SHEET_ID)
    return _sh.worksheet(MAIN_SHEET)

# ── 헬퍼 함수 ──────────────────────────────────────────────
def clean(v):
    if v is None: return ''
    if isinstance(v, float) and math.isnan(v): return ''
    s = str(v).strip()
    if s in ('nan','None','NaT'): return ''
    # 제어문자 제거 (줄바꿈은 \n으로 유지)
    s = ''.join(ch if ch == '\n' or ord(ch) >= 32 else ' ' for ch in s)
    # JSON 파싱 오류 유발하는 특수 따옴표 → 일반 따옴표로 변환
    s = s.replace('\u201c', "'").replace('\u201d', "'")  # "" → ''
    s = s.replace('\u2018', "'").replace('\u2019', "'")  # '' → ''
    s = s.replace('"', "'")  # 큰따옴표 → 작은따옴표
    return s

def calc_val_range(pre):
    """Pre 기업가치 기준으로 0~150, 150~360 자동 판별"""
    try:
        v = float(str(pre).replace(',','').replace('억원','').strip())
        v0_150 = 'O' if 0 < v <= 150 else 'X'
        v150_360 = 'O' if 150 < v <= 360 else 'X'
        return v0_150, v150_360
    except:
        return '', ''

# 컬럼 헤더 → 인덱스 매핑 (헤더행 기준)
COL_MAP = {
    'No.': 'A',
    '보고서 보관 위치': 'B',
    'GP명': 'C',
    '담당심사역': 'D',
    '펀드명': 'E',
    '사업자등록번호': 'F',
    '투자기업명': 'G',
    '업종(표준산업)': 'H',
    '주요사업(서비스)': 'I',
    '투자금액(억원)': 'J',
    '투자일자': 'K',
    '투자일자 기준': 'L',
    '기업가치(Pre, 억원)': 'M',
    '기업가치(Post, 억원) ': 'N',
    'Post : Outstanding(발행주식수 기준)/Full-dilution(완전희석 기준)': 'O',
    '투자검토보고서 링크': 'P',
    '비고': 'Q',
    '배치신청횟수': 'R',      # 읽기 전용
    '배치선정 기수': 'S',     # 읽기 전용
    '배치지원루트': 'T',      # 읽기 전용
    '직접투자 여부 (O/X)': 'U', # 읽기 전용
    '직접투자 일자': 'V',     # 읽기 전용
    '디캠프 기준 산업분류': 'W',
    '기업가치(0~150억원)': 'X',
    '기업가치(150~360억원)': 'Y',
    '자료 확인 여부': 'Z',
    '확인자': 'AA',
    '확인일자': 'AB',
    '기업 핵심 요약': 'AC',
    '투자포인트': 'AD',
    '주요리스크': 'AE',
    '기타 참고사항': 'AF',
    '검토 시 재확인 포인트': 'AG',
    '투자검토진행 의향': 'AH',
    '사유': 'AI',
    '담당자배정': 'AJ',
}

# 읽기 전용 컬럼 (웹에서 쓰기 금지)
READ_ONLY_COLS = {'배치신청횟수', '배치선정 기수', '배치지원루트', '직접투자 여부 (O/X)', '직접투자 일자'}

def col_letter_to_index(letter):
    """A→1, B→2, AA→27 변환"""
    result = 0
    for ch in letter.upper():
        result = result * 26 + (ord(ch) - ord('A') + 1)
    return result

# ── API 엔드포인트 ──────────────────────────────────────────

@app.route('/api/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok'})

@app.route('/api/data', methods=['GET'])
def get_data():
    """구글 시트 전체 데이터 읽기"""
    try:
        import sys, traceback
        print("=== /api/data 요청 시작 ===", flush=True)
        ws = get_sheet()
        print("=== 시트 연결 성공 ===", flush=True)
        all_values = ws.get_all_values()
        print(f"=== 전체 행 수: {len(all_values)} ===", flush=True)

        header_row_idx = HEADER_ROW - 1
        headers = all_values[header_row_idx]

        records = []
        for row in all_values[header_row_idx + 1:]:
            company_val = row[6].strip() if len(row) > 6 else ''
            if not company_val:
                continue
            no_val = row[0].strip() if row else ''
            if no_val == '-':
                continue

            record = {}
            for i, header in enumerate(headers):
                if not header:
                    continue
                val = row[i] if i < len(row) else ''
                record[header] = clean(val)

            pre = record.get('기업가치(Pre, 억원)', '')
            if pre and record.get('기업가치(0~150억원)', '') == '':
                v0_150, v150_360 = calc_val_range(pre)
                record['기업가치(0~150억원)'] = v0_150
                record['기업가치(150~360억원)'] = v150_360

            records.append(record)

        print(f"=== 읽은 레코드 수: {len(records)} ===", flush=True)

        import json
        safe_records = []
        for rec in records:
            try:
                json.dumps(rec, ensure_ascii=False)
                safe_records.append(rec)
            except Exception as e:
                print(f"JSON 오류 (No.{rec.get('No.','?')} {rec.get('투자기업명','?')}): {e}", flush=True)

        print(f"=== 최종 반환: {len(safe_records)}건 ===", flush=True)
        return jsonify({'success': True, 'data': safe_records, 'count': len(safe_records)})
    except Exception as e:
        import traceback
        print(f"=== 오류 발생: {str(e)} ===", flush=True)
        print(traceback.format_exc(), flush=True)
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/data/<row_no>', methods=['PUT'])
def update_row(row_no):
    """특정 행 업데이트 (No. 기준으로 행 찾기)"""
    try:
        body = request.json
        if not body:
            return jsonify({'success': False, 'error': '데이터 없음'}), 400

        ws = get_sheet()
        all_values = ws.get_all_values()
        header_row_idx = HEADER_ROW - 1
        headers = all_values[header_row_idx]

        # No. 컬럼으로 행 찾기
        target_row_idx = None
        for i, row in enumerate(all_values[header_row_idx + 1:], start=header_row_idx + 2):
            if row and str(row[0]).strip() == str(row_no):
                target_row_idx = i
                break

        if target_row_idx is None:
            return jsonify({'success': False, 'error': f'No.{row_no} 행을 찾을 수 없음'}), 404

        # 업데이트할 셀 목록 작성
        updates = []
        pre_val = None

        for field, value in body.items():
            # 읽기 전용 컬럼은 건너뜀
            if field in READ_ONLY_COLS:
                continue
            # 헤더에서 컬럼 인덱스 찾기
            if field in headers:
                col_idx = headers.index(field) + 1  # 1-based
                updates.append({'row': target_row_idx, 'col': col_idx, 'value': value})
                if field == '기업가치(Pre, 억원)':
                    pre_val = value

        # Pre 값이 바뀌면 X/Y열 자동 업데이트
        if pre_val is not None:
            v0_150, v150_360 = calc_val_range(pre_val)
            for field, val in [('기업가치(0~150억원)', v0_150), ('기업가치(150~360억원)', v150_360)]:
                if field in headers:
                    col_idx = headers.index(field) + 1
                    updates.append({'row': target_row_idx, 'col': col_idx, 'value': val})

        # 배치 업데이트 (API 호출 최소화)
        if updates:
            cells = [gspread.Cell(u['row'], u['col'], u['value']) for u in updates]
            ws.update_cells(cells, value_input_option='USER_ENTERED')

        return jsonify({'success': True, 'updated': len(updates), 'row': target_row_idx})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/data', methods=['POST'])
def add_row():
    """새 행 추가 (펀드팀 신규 입력)"""
    try:
        body = request.json
        if not body:
            return jsonify({'success': False, 'error': '데이터 없음'}), 400

        ws = get_sheet()
        all_values = ws.get_all_values()
        header_row_idx = HEADER_ROW - 1
        headers = all_values[header_row_idx]

        # 마지막 No. 찾아서 +1
        last_no = 0
        for row in all_values[header_row_idx + 1:]:
            if row and str(row[0]).strip().isdigit():
                last_no = max(last_no, int(row[0].strip()))
        new_no = last_no + 1

        # 새 행 데이터 구성
        new_row = [''] * len(headers)
        new_row[0] = str(new_no)  # No. 자동 부여

        for field, value in body.items():
            if field in READ_ONLY_COLS:
                continue
            if field in headers:
                col_idx = headers.index(field)
                new_row[col_idx] = value

        # Pre 기준 기업가치 구간 자동 계산
        pre_val = body.get('기업가치(Pre, 억원)', '')
        if pre_val:
            v0_150, v150_360 = calc_val_range(pre_val)
            for field, val in [('기업가치(0~150억원)', v0_150), ('기업가치(150~360억원)', v150_360)]:
                if field in headers:
                    new_row[headers.index(field)] = val

        # 데이터 마지막 행 다음에 추가
        ws.append_row(new_row, value_input_option='USER_ENTERED')

        return jsonify({'success': True, 'no': new_no})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/data/<row_no>', methods=['DELETE'])
def delete_row(row_no):
    """행 삭제"""
    try:
        ws = get_sheet()
        all_values = ws.get_all_values()
        header_row_idx = HEADER_ROW - 1

        target_row_idx = None
        for i, row in enumerate(all_values[header_row_idx + 1:], start=header_row_idx + 2):
            if row and str(row[0]).strip() == str(row_no):
                target_row_idx = i
                break

        if target_row_idx is None:
            return jsonify({'success': False, 'error': f'No.{row_no} 행을 찾을 수 없음'}), 404

        ws.delete_rows(target_row_idx)
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


if __name__ == '__main__':
    app.run(debug=True, port=5000)
