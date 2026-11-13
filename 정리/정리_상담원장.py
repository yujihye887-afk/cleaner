# 정리_상담원장.py — 상담원장_2025.xlsx 정리
#
# 결정 기록: my-wiki/decisions/d-002-상담원장-정리-다섯-결정.md
#   d-002-1  «종결»=«완료», «대기»=«보류» 로 합친다
#   d-002-2  처리시간 음수는 지우지 않고 값 그대로, 맨 끝 _표시 칸에 이유를 적는다
#   d-002-3  만족도 6~10점은 지우지 않고 값 그대로, _표시 칸에 이유를 적는다. 평균 만족도에서는 표시한 줄을 뺀다
#   d-002-4  칸 이름은 원본 그대로 (상담시간 칸도 «상담시간»)
#   d-002-5  고객번호 앞뒤 공백을 없앤다
# 위 다섯 말고는 이전 대화에서 정한 처리를 그대로 둔다 (결정 기록 없음).
#
# 원본은 읽기만 한다. 쓰는 것: 정리본/상담원장_정리.csv · 정리본/지운줄.csv · 정리본/변경요약.md
# 사용: python 정리/정리_상담원장.py   (어느 폴더에서 돌려도 된다)
import openpyxl, collections, re, csv, os

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(BASE, '..', 'messy', '상담원장_2025.xlsx')
OUT = os.path.join(BASE, '정리본')
os.makedirs(OUT, exist_ok=True)

rows = list(openpyxl.load_workbook(SRC, read_only=True)['원장'].iter_rows(values_only=True))
H = list(rows[0])
orig = [dict(zip(H, r)) for r in rows[1:]]
for i, x in enumerate(orig):
    x['_idx'] = i
N0 = len(orig)

nd = lambda s: re.sub(r'[/.]', '-', str(s).strip())
CH = {'전화': '전화', '유선': '전화', '전화상담': '전화', 'TEL': '전화', '챗봇': '챗봇', 'chatbot': '챗봇', 'Chatbot': '챗봇', '챗 봇': '챗봇',
      '앱': '앱', '모바일앱': '앱', 'app': '앱', 'APP': '앱', '이메일': '이메일', 'e-mail': '이메일', 'E-mail': '이메일', '메일': '이메일'}
REG = {'서울시': '서울', '서울특별시': '서울', '경기도': '경기'}
ST = {'완료': '완료', '처리완료': '완료', '종결': '완료', '진행중': '진행중', '처리중': '진행중', '보류': '보류', '대기': '보류'}   # d-002-1
Y = {'Y', 'y', '예', 'TRUE'}
N = {'N', 'n', '아니오', 'FALSE'}
FLAG_NEG = '처리시간 음수 — 틀린 값이지만 진짜 값을 몰라 그대로 둠 (d-002-2)'
FLAG_SAT = '만족도 6~10점 — NPS를 잘못 넣은 것으로 보이나 확인 못 해 그대로 둠, 평균 만족도에서 제외 (d-002-3)'

removed = []   # (원본 줄, 이유)
# 1 테스트 접수
rest = []
for x in orig:
    if x['상담원ID'] == 'AG999':
        removed.append((x, '테스트 접수(상담원ID AG999)'))
    else:
        rest.append(x)
# 2 같은 접수번호 중복: 최종수정일 늦은 줄만 남김
g = collections.defaultdict(list)
for x in rest:
    g[x['접수번호']].append(x)
kept = []
for k, v in g.items():
    v = sorted(v, key=lambda y: nd(y['최종수정일']))   # 안정 정렬: 같은 날짜면 원본에서 뒤에 있는 줄이 남음
    top = v[-1]
    kept.append(top)
    for y in v[:-1]:
        removed.append((y, '중복 추출(최종수정일이 같은 같은 줄)' if nd(y['최종수정일']) == nd(top['최종수정일'])
                        else '같은 접수번호의 옛 상태 줄(최종수정일이 더 이른 줄)'))

out = []
sec_lines = 0
for x in kept:
    o = {c: x[c] for c in H}
    for k in ('접수일', '최종수정일', '가입일'):
        if o[k] is not None:
            o[k] = nd(o[k])
    o['고객번호'] = o['고객번호'].strip()                       # d-002-5
    o['접수채널'] = CH[o['접수채널'].strip()]
    if o['지역']:
        o['지역'] = REG.get(o['지역'], o['지역'])
    if isinstance(o['월정액'], str):
        o['월정액'] = int(re.sub(r'\D', '', o['월정액']))
    o['상담원팀'] = o['상담원팀'].replace(' ', '')
    o['처리상태'] = ST[o['처리상태'].strip()]
    r = o['재문의']
    o['재문의'] = 1 if r in Y else 0 if r in N else None
    m = o['만족도']
    if isinstance(m, str):
        m = m.replace('점', '').strip()
    try:
        m = int(m)
    except (ValueError, TypeError):
        m = None
    flag = ''
    if m is not None and m < 1:
        m = None                                                # 0 · «-» · «미응답» → 빈칸
    if m is not None and m > 5:
        flag = FLAG_SAT                                         # d-002-3: 값은 그대로
    o['만족도'] = m
    s = o['상담시간']
    if s >= 9999:
        s = None
    elif o['접수채널'] == '전화' and o['접수일'] >= '2025-07-01':
        s = s / 60
        sec_lines += 1
    o['상담시간'] = None if s is None else round(s, 2)           # d-002-4: 칸 이름은 «상담시간» 그대로, 값은 분
    if o['처리시간'] is not None and o['처리시간'] < 0:
        flag = FLAG_NEG                                         # d-002-2: 값은 그대로
    o['_표시'] = flag
    o['_원본'] = x
    out.append(o)
out.sort(key=lambda o: (o['접수일'], o['접수시각'], o['접수번호']))

COLS = H + ['_표시']
with open(os.path.join(OUT, '상담원장_정리.csv'), 'w', encoding='utf-8-sig', newline='') as f:
    w = csv.writer(f)
    w.writerow(COLS)
    for o in out:
        w.writerow(['' if o[c] is None else o[c] for c in COLS])
removed.sort(key=lambda t: t[0]['_idx'])
with open(os.path.join(OUT, '지운줄.csv'), 'w', encoding='utf-8-sig', newline='') as f:
    w = csv.writer(f)
    w.writerow(H + ['지운이유'])
    for x, why in removed:
        w.writerow(['' if x[c] is None else x[c] for c in H] + [why])


# ---- 변경요약.md
def same(a, b):
    return ('' if a is None else str(a)) == ('' if b is None else str(b))


chg = {c: sum(1 for o in out if not same(o['_원본'][c], o[c])) for c in H}
rc = collections.Counter(why for _, why in removed)
n = len(out)
flag_neg = sum(1 for o in out if o['_표시'] == FLAG_NEG)
flag_sat = sum(1 for o in out if o['_표시'] == FLAG_SAT)
orig_st = collections.Counter(str(o['_원본']['처리상태']).strip() for o in out)
L = []
L.append('# 정리 변경 요약\n')
L.append('원본 `상담원장_2025.xlsx`(시트 «원장»)는 읽기만 했다. 결정 기록: `my-wiki/decisions/d-002-상담원장-정리-다섯-결정.md`\n')
L.append('## 줄 수\n')
L.append('| 단계 | 줄 수 |\n|---|---|')
L.append(f'| 원본 | {N0:,} |')
for k in ('테스트 접수(상담원ID AG999)', '같은 접수번호의 옛 상태 줄(최종수정일이 더 이른 줄)', '중복 추출(최종수정일이 같은 같은 줄)'):
    L.append(f'| − {k} | {rc[k]:,} |')
L.append(f'| 정리본 | {n:,} |\n')
L.append('## 상담시간 칸 — 초에서 분으로\n')
L.append(f'- 칸 이름은 원본 그대로 «상담시간»이다. **2025-07-01 이후 접수된 전화 채널 {sec_lines:,}줄은 값이 초에서 분으로 바뀌었다 (60으로 나눔, 소수 둘째 자리).** 나머지 줄은 원래 분 단위라 그대로다.')
L.append('- 전화 채널은 «전화»·«유선»·«전화상담»·«TEL»을 합친 것이다.\n')
L.append('## 칸마다 값이 바뀐 줄 수 (정리본 줄 기준)\n')
L.append('| 칸 | 바뀐 줄 |\n|---|---|')
for c in H:
    if chg[c]:
        L.append(f'| {c} | {chg[c]:,} |')
L.append('')
L.append('칸 이름은 하나도 바꾸지 않았다. 맨 끝에 `_표시` 칸 하나를 더했다 (24칸 → 25칸).\n')
L.append('## 결정 기록대로 처리한 것\n')
L.append('| 결정 | 처리 | 줄 수 |\n|---|---|---|')
L.append(f'| d-002-1 | «종결» → «완료» | {orig_st["종결"]:,} |')
L.append(f'| d-002-1 | «대기» → «보류» | {orig_st["대기"]:,} |')
L.append(f'| d-002-2 | 처리시간 음수: 값 그대로, `_표시`에 이유 | {flag_neg:,} |')
L.append(f'| d-002-3 | 만족도 6~10점: 값 그대로, `_표시`에 이유 (평균 만족도에서 제외) | {flag_sat:,} |')
L.append('| d-002-4 | 칸 이름 원본 그대로 | 바뀐 칸 0개 |')
L.append(f'| d-002-5 | 고객번호 앞뒤 공백 제거 | {chg["고객번호"]:,} |\n')
L.append('## 그 밖에 바꾼 것 (이전 대화에서 정한 처리, 결정 기록 없음)\n')
L.append(f'- 날짜 표기를 `YYYY-MM-DD`로 통일: 접수일 {chg["접수일"]:,} · 최종수정일 {chg["최종수정일"]:,} · 가입일 {chg["가입일"]:,}줄')
L.append(f'- 접수채널 표기 통일 {chg["접수채널"]:,}줄 · 지역 {chg["지역"]:,}줄 · 월정액 {chg["월정액"]:,}줄 · 상담원팀 공백 {chg["상담원팀"]:,}줄')
L.append(f'- 재문의를 1/0으로: {chg["재문의"]:,}줄 (빈칸은 빈칸)')
L.append(f'- 만족도: 값이 바뀐 줄 {chg["만족도"]:,} («5점»→5, 0·«-»·«미응답»→빈칸)')
L.append(f'- 처리상태: 값이 바뀐 줄 {chg["처리상태"]:,} («처리완료»·«종결»·«완료 »→«완료», «처리중»→«진행중», «대기»→«보류»)')
with open(os.path.join(OUT, '변경요약.md'), 'w', encoding='utf-8', newline='\n') as f:
    f.write('\n'.join(L) + '\n')

# ---- 화면
tm = [o['상담시간'] for o in out if o['상담시간'] is not None]
rq = [o['재문의'] for o in out if o['재문의'] is not None]
sa = [o['만족도'] for o in out if o['만족도'] is not None and o['_표시'] != FLAG_SAT]   # d-002-3: 6~10점으로 표시한 줄을 뺀다
cp = sum(1 for o in out if o['처리상태'] == '완료')
print(f'줄: 원본 {N0} → 정리본 {n} (지운 줄 {len(removed)})')
print(f'상담 건수 {n} | 평균 상담시간 {sum(tm)/len(tm):.2f}분 | 재문의율 {100*sum(rq)/len(rq):.2f}% ({sum(rq)}/{len(rq)}) | 평균 만족도 {sum(sa)/len(sa):.3f} (응답 {len(sa)}) | 완료율 {100*cp/n:.2f}% ({cp}/{n})')
print(f'고객 수 {len({o["고객번호"] for o in out})}')
