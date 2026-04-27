import os, json, math
from flask import Flask, jsonify, request
from flask_cors import CORS
import gspread
from google.oauth2.service_account import Credentials

app = Flask(__name__)
CORS(app, origins=[
    'https://gayoung-dcamp.github.io',
    'http://localhost:3000',
    'http://127.0.0.1:3000',
    'null'  # 로컬 파일 접근
], supports_credentials=False)

# ── 구글 시트 연결 ──────────────────────────────────────────
SHEET_ID = '15AnatVs4sauXt2FXLkqVzpmyTtVdpGRNTjmVx7VKoZ0'
MAIN_SHEET = '펀드 투자검토보고서 업데이트 현황_20260203'
POLICY_SHEET = '집중투자_정책_연도별'  # 신규: 정책 시트
HEADER_ROW = 10  # 메인 시트의 헤더 행

SCOPES = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']

_gc = None
_sh = None

def get_sheet(sheet_name=MAIN_SHEET):
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
    return _sh.worksheet(sheet_name)

# ── 헬퍼 함수 ──────────────────────────────────────────────
def clean(v):
    if v is None: return ''
    if isinstance(v, float) and math.isnan(v): return ''
    s = str(v).strip()
    if s in ('nan','None','NaT'): return ''
    # 제어문자 제거 (줄바꿈은 \n으로 유지)
    s = ''.join(ch if ch == '\n' or ord(ch) >= 32 else ' ' for ch in s)
    # JSON 파싱 오류 유발하는 특수 따옴표 → 일반 따옴표로 변환
    s = s.replace('\u201c', "'").replace('\u201d', "'")
    s = s.replace('\u2018', "'").replace('\u2019', "'")
    s = s.replace('"', "'")
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

# 컬럼 헤더 → 인덱스 매핑 (참조용 — 실제 동작은 헤더 검색 사용)
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
    '배치신청횟수': 'R',
    '배치선정 기수': 'S',
    '배치지원루트': 'T',
    '직접투자 여부 (O/X)': 'U',
    '직접투자 일자': 'V',
    '디캠프 기준 산업분류': 'W',
    '집중투자연도': 'X',  # 신규 추가
    '기업가치(0~150억원)': 'Y',
    '기업가치(150~360억원)': 'Z',
    '자료 확인 여부': 'AA',
    '확인자': 'AB',
    '확인일자': 'AC',
    '기업 핵심 요약': 'AD',
    '투자포인트': 'AE',
    '주요리스크': 'AF',
    '기타 참고사항': 'AG',
    '검토 시 재확인 포인트': 'AH',
    '투자검토진행 의향': 'AI',
    '사유': 'AJ',
    '담당자배정': 'AK',
}

# 읽기 전용 컬럼 (웹에서 쓰기 금지)
READ_ONLY_COLS = {'배치신청횟수', '배치선정 기수', '배치지원루트', '직접투자 여부 (O/X)', '직접투자 일자'}

def col_letter_to_index(letter):
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

        target_row_idx = None
        for i, row in enumerate(all_values[header_row_idx + 1:], start=header_row_idx + 2):
            if row and str(row[0]).strip() == str(row_no):
                target_row_idx = i
                break

        if target_row_idx is None:
            return jsonify({'success': False, 'error': f'No.{row_no} 행을 찾을 수 없음'}), 404

        updates = []
        pre_val = None
        user_sent_0_150 = '기업가치(0~150억원)' in body
        user_sent_150_360 = '기업가치(150~360억원)' in body

        for field, value in body.items():
            if field in READ_ONLY_COLS:
                continue
            if field in headers:
                col_idx = headers.index(field) + 1
                updates.append({'row': target_row_idx, 'col': col_idx, 'value': value})
                if field == '기업가치(Pre, 억원)':
                    pre_val = value

        if pre_val is not None:
            v0_150, v150_360 = calc_val_range(pre_val)
            auto_updates = []
            if not user_sent_0_150:
                auto_updates.append(('기업가치(0~150억원)', v0_150))
            if not user_sent_150_360:
                auto_updates.append(('기업가치(150~360억원)', v150_360))
            for field, val in auto_updates:
                if field in headers:
                    col_idx = headers.index(field) + 1
                    updates.append({'row': target_row_idx, 'col': col_idx, 'value': val})

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

        last_no = 0
        for row in all_values[header_row_idx + 1:]:
            if row and str(row[0]).strip().isdigit():
                last_no = max(last_no, int(row[0].strip()))
        new_no = last_no + 1

        new_row = [''] * len(headers)
        new_row[0] = str(new_no)

        for field, value in body.items():
            if field in READ_ONLY_COLS:
                continue
            if field in headers:
                col_idx = headers.index(field)
                new_row[col_idx] = value

        pre_val = body.get('기업가치(Pre, 억원)', '')
        user_sent_0_150 = body.get('기업가치(0~150억원)', '').strip() != ''
        user_sent_150_360 = body.get('기업가치(150~360억원)', '').strip() != ''
        if pre_val:
            v0_150, v150_360 = calc_val_range(pre_val)
            if not user_sent_0_150 and '기업가치(0~150억원)' in headers:
                new_row[headers.index('기업가치(0~150억원)')] = v0_150
            if not user_sent_150_360 and '기업가치(150~360억원)' in headers:
                new_row[headers.index('기업가치(150~360억원)')] = v150_360

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


# ════════════════════════════════════════════════════════════
# 신규: 집중투자 정책 관리 API
# ════════════════════════════════════════════════════════════

@app.route('/api/focus-policy', methods=['GET'])
def get_focus_policy():
    """집중투자_정책_연도별 시트 전체 읽기"""
    try:
        ws = get_sheet(POLICY_SHEET)
        all_values = ws.get_all_values()

        if len(all_values) < 2:
            return jsonify({'success': True, 'policies': {}, 'years': []})

        # 1행 = 헤더, 2행부터 데이터
        # A: 연도, B: 집중투자 산업 (콤마 구분), C: 메모
        policies = {}
        years = []
        for row in all_values[1:]:
            if not row or not row[0].strip():
                continue
            year_str = clean(row[0]).strip()
            if not year_str:
                continue

            industries_raw = clean(row[1]) if len(row) > 1 else ''
            memo = clean(row[2]) if len(row) > 2 else ''

            # 콤마로 분리, 공백 제거, 빈 값 제외
            industries = [s.strip() for s in industries_raw.split(',') if s.strip()]

            policies[year_str] = {
                'year': year_str,
                'industries': industries,
                'memo': memo
            }
            years.append(year_str)

        # 연도 정렬 (숫자로 변환 가능한 것만)
        try:
            years.sort(key=lambda x: int(x))
        except ValueError:
            years.sort()

        return jsonify({
            'success': True,
            'policies': policies,
            'years': years,
            'current_year': years[-1] if years else None
        })
    except Exception as e:
        import traceback
        print(f"=== 정책 조회 오류: {str(e)} ===", flush=True)
        print(traceback.format_exc(), flush=True)
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/focus-policy/<year>', methods=['PUT'])
def update_focus_policy(year):
    """특정 연도의 정책 업데이트 (없으면 추가)"""
    try:
        body = request.json
        if not body:
            return jsonify({'success': False, 'error': '데이터 없음'}), 400

        industries = body.get('industries', [])
        memo = body.get('memo', '')

        # industries는 리스트로 받아서 콤마 결합
        if isinstance(industries, list):
            industries_str = ','.join([s.strip() for s in industries if s and s.strip()])
        else:
            industries_str = str(industries).strip()

        ws = get_sheet(POLICY_SHEET)
        all_values = ws.get_all_values()

        # 해당 연도 행 찾기
        target_row = None
        for i, row in enumerate(all_values[1:], start=2):
            if row and clean(row[0]).strip() == str(year).strip():
                target_row = i
                break

        if target_row is None:
            # 없으면 새 행 추가
            ws.append_row([str(year), industries_str, memo], value_input_option='USER_ENTERED')
            return jsonify({'success': True, 'action': 'added', 'year': year})
        else:
            # 있으면 업데이트
            cells = [
                gspread.Cell(target_row, 1, str(year)),
                gspread.Cell(target_row, 2, industries_str),
                gspread.Cell(target_row, 3, memo),
            ]
            ws.update_cells(cells, value_input_option='USER_ENTERED')
            return jsonify({'success': True, 'action': 'updated', 'year': year})

    except Exception as e:
        import traceback
        print(f"=== 정책 업데이트 오류: {str(e)} ===", flush=True)
        print(traceback.format_exc(), flush=True)
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/focus-policy/<year>', methods=['DELETE'])
def delete_focus_policy(year):
    """특정 연도 정책 삭제"""
    try:
        ws = get_sheet(POLICY_SHEET)
        all_values = ws.get_all_values()

        target_row = None
        for i, row in enumerate(all_values[1:], start=2):
            if row and clean(row[0]).strip() == str(year).strip():
                target_row = i
                break

        if target_row is None:
            return jsonify({'success': False, 'error': f'{year}년 정책을 찾을 수 없음'}), 404

        ws.delete_rows(target_row)
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


if __name__ == '__main__':
    app.run(debug=True, port=5000)
