# -*- coding: utf-8 -*-
"""Đấu trường cờ vua — bộ luật cờ vua phía máy chủ (gương của cv-engine.js)
và logic ván đấu / giải đấu trên dict thuần (không phụ thuộc FastAPI;
main.py lo file IO, khóa và HTTP).

Bàn: list 64 ô, index = hàng*8 + cột, hàng 0 = hàng 8 (đen).
Quân: trắng 'K','Q','R','B','N','P'; đen 'k','q','r','b','n','p'; None = trống.
"""
import time

KN = [(-2, -1), (-2, 1), (-1, -2), (-1, 2), (1, -2), (1, 2), (2, -1), (2, 1)]
KG = [(-1, -1), (-1, 0), (-1, 1), (0, -1), (0, 1), (1, -1), (1, 0), (1, 1)]
DIAG = [(-1, -1), (-1, 1), (1, -1), (1, 1)]
ORTH = [(-1, 0), (1, 0), (0, -1), (0, 1)]


def _rc(i):
    return divmod(i, 8)


def _ix(r, c):
    return r * 8 + c


def _inside(r, c):
    return 0 <= r < 8 and 0 <= c < 8


def _white(p):
    return p is not None and p == p.upper()


def _black(p):
    return p is not None and p == p.lower()


def _same(a, b):
    return (_white(a) and _white(b)) or (_black(a) and _black(b))


def start_state():
    back = ['r', 'n', 'b', 'q', 'k', 'b', 'n', 'r']
    b = [None] * 64
    for c in range(8):
        b[_ix(0, c)] = back[c]
        b[_ix(1, c)] = 'p'
        b[_ix(6, c)] = 'P'
        b[_ix(7, c)] = back[c].upper()
    return {'b': b, 'turn': 'w', 'castling': {'K': True, 'Q': True, 'k': True, 'q': True},
            'ep': None, 'half': 0, 'full': 1}


def parse_fen(fen):
    parts = fen.strip().split()
    rows = parts[0].split('/')
    b = [None] * 64
    for r in range(8):
        c = 0
        for ch in rows[r]:
            if ch.isdigit():
                c += int(ch)
            else:
                b[_ix(r, c)] = ch
                c += 1
    cast = {'K': False, 'Q': False, 'k': False, 'q': False}
    if len(parts) > 2 and parts[2] != '-':
        for k in parts[2]:
            cast[k] = True
    ep = None
    if len(parts) > 3 and parts[3] != '-':
        ep = _ix(8 - int(parts[3][1]), ord(parts[3][0]) - 97)
    return {'b': b, 'turn': parts[1] if len(parts) > 1 else 'w', 'castling': cast, 'ep': ep,
            'half': int(parts[4]) if len(parts) > 4 else 0,
            'full': int(parts[5]) if len(parts) > 5 else 1}


def board_str(b):
    return ''.join(p or '.' for p in b)


def state_from_str(s, turn='w', castling='', ep=-1, half=0, full=1):
    b = [None if ch == '.' else ch for ch in s]
    cast = {'K': False, 'Q': False, 'k': False, 'q': False}
    for k in (castling or ''):
        if k in cast:
            cast[k] = True
    return {'b': b, 'turn': turn, 'castling': cast,
            'ep': (None if ep in (None, -1) else ep), 'half': half, 'full': full}


def castling_str(cast):
    s = ''.join(k for k in 'KQkq' if cast.get(k))
    return s or '-'


def pos_key(st):
    return board_str(st['b']) + st['turn'] + castling_str(st['castling']) + str(st['ep'])


def find_king(st, white):
    t = 'K' if white else 'k'
    for i, p in enumerate(st['b']):
        if p == t:
            return i
    return -1


def attacked(st, sq, by_white):
    r, c = _rc(sq)
    pr = r + 1 if by_white else r - 1
    pp = 'P' if by_white else 'p'
    for dc in (-1, 1):
        if _inside(pr, c + dc) and st['b'][_ix(pr, c + dc)] == pp:
            return True
    np_ = 'N' if by_white else 'n'
    for dr, dc in KN:
        rr, cc = r + dr, c + dc
        if _inside(rr, cc) and st['b'][_ix(rr, cc)] == np_:
            return True
    kp = 'K' if by_white else 'k'
    for dr, dc in KG:
        rr, cc = r + dr, c + dc
        if _inside(rr, cc) and st['b'][_ix(rr, cc)] == kp:
            return True
    for dr, dc in DIAG:
        rr, cc = r + dr, c + dc
        while _inside(rr, cc):
            p = st['b'][_ix(rr, cc)]
            if p:
                if (by_white and p in ('B', 'Q')) or ((not by_white) and p in ('b', 'q')):
                    return True
                break
            rr += dr
            cc += dc
    for dr, dc in ORTH:
        rr, cc = r + dr, c + dc
        while _inside(rr, cc):
            p = st['b'][_ix(rr, cc)]
            if p:
                if (by_white and p in ('R', 'Q')) or ((not by_white) and p in ('r', 'q')):
                    return True
                break
            rr += dr
            cc += dc
    return False


def in_check(st, white):
    k = find_king(st, white)
    return k >= 0 and attacked(st, k, not white)


def gen_moves(st):
    white = st['turn'] == 'w'
    moves = []

    def push(f, t, **kw):
        m = {'from': f, 'to': t}
        m.update(kw)
        moves.append(m)

    for i, p in enumerate(st['b']):
        if not p or (white != _white(p)):
            continue
        r, c = _rc(i)
        up = p.upper()
        if up == 'P':
            d = -1 if white else 1
            sr = 6 if white else 1
            lr = 0 if white else 7
            r1 = r + d
            if _inside(r1, c) and not st['b'][_ix(r1, c)]:
                if r1 == lr:
                    for pr in ('Q', 'R', 'B', 'N'):
                        push(i, _ix(r1, c), promo=(pr if white else pr.lower()))
                else:
                    push(i, _ix(r1, c))
                    r2 = r + 2 * d
                    if r == sr and not st['b'][_ix(r2, c)]:
                        push(i, _ix(r2, c), dbl=True)
            for dc in (-1, 1):
                cc = c + dc
                if not _inside(r1, cc):
                    continue
                t = st['b'][_ix(r1, cc)]
                if t and not _same(p, t):
                    if r1 == lr:
                        for pr in ('Q', 'R', 'B', 'N'):
                            push(i, _ix(r1, cc), promo=(pr if white else pr.lower()))
                    else:
                        push(i, _ix(r1, cc))
                elif (not t) and st['ep'] == _ix(r1, cc):
                    push(i, _ix(r1, cc), ep=True)
        elif up in ('N', 'K'):
            steps = KN if up == 'N' else KG
            for dr, dc in steps:
                rr, cc = r + dr, c + dc
                if not _inside(rr, cc):
                    continue
                t = st['b'][_ix(rr, cc)]
                if (not t) or (not _same(p, t)):
                    push(i, _ix(rr, cc))
            if up == 'K':
                home = 7 if white else 0
                rk = ('K', 'Q') if white else ('k', 'q')
                if i == _ix(home, 4) and not in_check(st, white):
                    if (st['castling'][rk[0]] and not st['b'][_ix(home, 5)] and not st['b'][_ix(home, 6)]
                            and st['b'][_ix(home, 7)] == ('R' if white else 'r')
                            and not attacked(st, _ix(home, 5), not white)
                            and not attacked(st, _ix(home, 6), not white)):
                        push(i, _ix(home, 6), castle='K')
                    if (st['castling'][rk[1]] and not st['b'][_ix(home, 3)] and not st['b'][_ix(home, 2)]
                            and not st['b'][_ix(home, 1)] and st['b'][_ix(home, 0)] == ('R' if white else 'r')
                            and not attacked(st, _ix(home, 3), not white)
                            and not attacked(st, _ix(home, 2), not white)):
                        push(i, _ix(home, 2), castle='Q')
        else:
            dirs = []
            if up in ('B', 'Q'):
                dirs += DIAG
            if up in ('R', 'Q'):
                dirs += ORTH
            for dr, dc in dirs:
                rr, cc = r + dr, c + dc
                while _inside(rr, cc):
                    t = st['b'][_ix(rr, cc)]
                    if not t:
                        push(i, _ix(rr, cc))
                    else:
                        if not _same(p, t):
                            push(i, _ix(rr, cc))
                        break
                    rr += dr
                    cc += dc
    legal = []
    for m in moves:
        ns = apply_move(st, m)
        if not in_check(ns, white):
            legal.append(m)
    return legal


def apply_move(st, m):
    b = list(st['b'])
    p = b[m['from']]
    castling = dict(st['castling'])
    fr, fc = _rc(m['from'])
    tr, tc = _rc(m['to'])
    captured = b[m['to']]
    if m.get('ep'):
        captured = b[_ix(fr, tc)]
        b[_ix(fr, tc)] = None
    b[m['to']] = m.get('promo') or p
    b[m['from']] = None
    if m.get('castle') == 'K':
        b[_ix(fr, 5)] = b[_ix(fr, 7)]
        b[_ix(fr, 7)] = None
    if m.get('castle') == 'Q':
        b[_ix(fr, 3)] = b[_ix(fr, 0)]
        b[_ix(fr, 0)] = None
    if p == 'K':
        castling['K'] = castling['Q'] = False
    if p == 'k':
        castling['k'] = castling['q'] = False
    if p == 'R' and m['from'] == _ix(7, 7):
        castling['K'] = False
    if p == 'R' and m['from'] == _ix(7, 0):
        castling['Q'] = False
    if p == 'r' and m['from'] == _ix(0, 7):
        castling['k'] = False
    if p == 'r' and m['from'] == _ix(0, 0):
        castling['q'] = False
    for sq, k in ((_ix(7, 7), 'K'), (_ix(7, 0), 'Q'), (_ix(0, 7), 'k'), (_ix(0, 0), 'q')):
        if m['to'] == sq:
            castling[k] = False
    ep = _ix((fr + tr) // 2, fc) if m.get('dbl') else None
    half = 0 if (p.upper() == 'P' or captured) else st['half'] + 1
    return {'b': b, 'turn': 'b' if st['turn'] == 'w' else 'w', 'castling': castling, 'ep': ep,
            'half': half, 'full': st['full'] + 1 if st['turn'] == 'b' else st['full']}


def insufficient_material(st):
    pcs = [p for p in st['b'] if p and p.upper() != 'K']
    if not pcs:
        return True
    if len(pcs) == 1:
        return pcs[0].upper() in ('B', 'N')
    if len(pcs) == 2 and pcs[0].upper() == 'B' and pcs[1].upper() == 'B':
        sq = [i for i, p in enumerate(st['b']) if p and p.upper() == 'B']
        c1 = (sq[0] // 8 + sq[0] % 8) % 2
        c2 = (sq[1] // 8 + sq[1] % 8) % 2
        if _white(st['b'][sq[0]]) != _white(st['b'][sq[1]]) and c1 == c2:
            return True
    return False


def status(st):
    moves = gen_moves(st)
    if not moves:
        return 'checkmate' if in_check(st, st['turn'] == 'w') else 'stalemate'
    if insufficient_material(st):
        return 'draw-insufficient'
    if st['half'] >= 100:
        return 'draw-fifty'
    return 'playing'


def sq_name(i):
    r, c = _rc(i)
    return 'abcdefgh'[c] + str(8 - r)


def move_san(st, m):
    if m.get('castle') == 'K':
        return 'O-O'
    if m.get('castle') == 'Q':
        return 'O-O-O'
    p = st['b'][m['from']]
    up = p.upper()
    cap = bool(st['b'][m['to']] or m.get('ep'))
    dest = sq_name(m['to'])
    if up == 'P':
        san = (('abcdefgh'[_rc(m['from'])[1]] + 'x') if cap else '') + dest
        if m.get('promo'):
            san += '=' + m['promo'].upper()
    else:
        letter = {'N': 'N', 'B': 'B', 'R': 'R', 'Q': 'Q', 'K': 'K'}[up]
        others = [x for x in gen_moves(st)
                  if x['from'] != m['from'] and x['to'] == m['to'] and st['b'][x['from']] == p]
        dis = ''
        if others:
            same_file = any(_rc(x['from'])[1] == _rc(m['from'])[1] for x in others)
            same_rank = any(_rc(x['from'])[0] == _rc(m['from'])[0] for x in others)
            if not same_file:
                dis = 'abcdefgh'[_rc(m['from'])[1]]
            elif not same_rank:
                dis = str(8 - _rc(m['from'])[0])
            else:
                dis = sq_name(m['from'])
        san = letter + dis + ('x' if cap else '') + dest
    ns = apply_move(st, m)
    opp = gen_moves(ns)
    if not opp and in_check(ns, ns['turn'] == 'w'):
        san += '#'
    elif in_check(ns, ns['turn'] == 'w'):
        san += '+'
    return san


def perft(st, depth):
    if depth == 0:
        return 1
    return sum(perft(apply_move(st, m), depth - 1) for m in gen_moves(st))


# ---------------- Logic ván đấu (dict thuần) ----------------

def game_state(g):
    return state_from_str(g['board'], g['turn'], g['castling'], g['ep'], g['half'], g['full'])


def game_sync(g, st):
    g['board'] = board_str(st['b'])
    g['turn'] = st['turn']
    g['castling'] = castling_str(st['castling'])
    g['ep'] = -1 if st['ep'] is None else st['ep']
    g['half'] = st['half']
    g['full'] = st['full']


def new_game(gid, white, white_name, black, black_name, mode, clock_min,
             room_code=None, tour_id=None, tour_match=None, inc=0, theme='classic', ranked=True):
    st = start_state()
    ms = int(clock_min) * 60000
    g = {'id': gid, 'room_code': room_code, 'tour_id': tour_id, 'tour_match': tour_match,
         'white': white, 'white_name': white_name, 'black': black, 'black_name': black_name,
         'mode': mode, 'clock_min': int(clock_min), 'inc': int(inc), 'theme': theme, 'ranked': bool(ranked),
         'status': 'playing', 'result': None,
         'reason': None, 'winner': None, 'draw_offer': None, 'moves': [],
         'positions': [pos_key(st)], 'clocks': {'w': ms, 'b': ms}, 'tick': time.time(),
         'created': time.time(), 'updated': time.time()}
    game_sync(g, st)
    return g


def live_clocks(g, now=None):
    now = now if now is not None else time.time()
    w, b = int(g['clocks']['w']), int(g['clocks']['b'])
    if g['status'] == 'playing':
        el = int((now - float(g['tick'])) * 1000)
        if g['turn'] == 'w':
            w = max(0, w - el)
        else:
            b = max(0, b - el)
    return {'w': w, 'b': b}


def finish_game(g, result, reason, winner=None):
    g['status'] = 'finished'
    g['result'] = result
    g['reason'] = reason
    g['winner'] = winner
    g['updated'] = time.time()
    return g


def apply_player_move(g, key, frm, to, promo=None, now=None):
    """Trả (game, lỗi-hoặc-None). Kiểm tra lượt, người chơi, tính hợp lệ, trạng thái."""
    now = now if now is not None else time.time()
    if g['status'] != 'playing':
        return g, 'Ván này đã kết thúc'
    side = 'w' if g['white'] == key else ('b' if g['black'] == key else None)
    if side is None:
        return g, 'Thầy cô không phải kỳ thủ của ván này'
    if side != g['turn']:
        return g, 'Chưa tới lượt của thầy cô'
    st = game_state(g)
    want_promo = None
    if promo:
        want_promo = promo.upper() if side == 'w' else promo.lower()
        if want_promo.upper() not in ('Q', 'R', 'B', 'N'):
            return g, 'Quân phong cấp không hợp lệ'
    mv = None
    for m in gen_moves(st):
        if m['from'] == frm and m['to'] == to and (m.get('promo') or None) == want_promo:
            mv = m
            break
    if mv is None:
        # Tốt tới cuối mà không chọn quân -> mặc định phong Hậu
        p = st['b'][frm]
        if p and p.upper() == 'P' and _rc(to)[0] in (0, 7):
            for m in gen_moves(st):
                if m['from'] == frm and m['to'] == to and m.get('promo') == ('Q' if side == 'w' else 'q'):
                    mv = m
                    break
    if mv is None:
        return g, 'Nước đi không hợp lệ'
    # Trừ đồng hồ bên vừa đi rồi cộng increment (nếu có)
    live = live_clocks(g, now)
    live[side] = live[side] + int(g.get('inc') or 0) * 1000
    g['clocks'] = live
    g['tick'] = now
    san = move_san(st, mv)
    ns = apply_move(st, mv)
    game_sync(g, ns)
    g['moves'].append({'from': frm, 'to': to, 'promo': mv.get('promo'), 'san': san,
                       'by': key, 'at': now})
    key_pos = pos_key(ns)
    g['positions'].append(key_pos)
    g['draw_offer'] = None
    g['updated'] = now
    stt = status(ns)
    if stt == 'checkmate':
        finish_game(g, '1-0' if side == 'w' else '0-1', 'checkmate',
                    g['white'] if side == 'w' else g['black'])
    elif stt == 'stalemate':
        finish_game(g, '1/2-1/2', 'stalemate')
    elif stt == 'draw-insufficient':
        finish_game(g, '1/2-1/2', 'insufficient')
    elif stt == 'draw-fifty':
        finish_game(g, '1/2-1/2', 'fifty')
    elif g['positions'].count(key_pos) >= 3:
        finish_game(g, '1/2-1/2', 'threefold')
    return g, None


def resign_game(g, key):
    if g['status'] != 'playing':
        return g, 'Ván này đã kết thúc'
    if key == g['white']:
        finish_game(g, '0-1', 'resign', g['black'])
    elif key == g['black']:
        finish_game(g, '1-0', 'resign', g['white'])
    else:
        return g, 'Thầy cô không phải kỳ thủ của ván này'
    return g, None


def draw_offer(g, key):
    if g['status'] != 'playing':
        return g, 'Ván này đã kết thúc'
    if key not in (g['white'], g['black']):
        return g, 'Thầy cô không phải kỳ thủ của ván này'
    if g.get('draw_offer') == key:
        return g, 'Thầy cô đã gửi đề nghị hòa rồi'
    g['draw_offer'] = key
    g['updated'] = time.time()
    return g, None


def draw_respond(g, key, accept):
    if g['status'] != 'playing':
        return g, 'Ván này đã kết thúc'
    offer = g.get('draw_offer')
    if not offer or offer == key:
        return g, 'Không có đề nghị hòa nào đang chờ thầy cô'
    if key not in (g['white'], g['black']):
        return g, 'Thầy cô không phải kỳ thủ của ván này'
    if accept:
        finish_game(g, '1/2-1/2', 'agreement')
    else:
        g['draw_offer'] = None
        g['updated'] = time.time()
    return g, None


def claim_timeout(g, key, now=None):
    now = now if now is not None else time.time()
    if g['status'] != 'playing':
        return g, 'Ván này đã kết thúc'
    if key not in (g['white'], g['black']):
        return g, 'Thầy cô không phải kỳ thủ của ván này'
    live = live_clocks(g, now)
    loser = None
    if live['w'] <= 0:
        loser = 'w'
    elif live['b'] <= 0:
        loser = 'b'
    if loser is None:
        return g, 'Đồng hồ vẫn còn giờ'
    st = game_state(g)
    # Đối thủ không đủ lực chiếu hết -> hòa
    opp_white = loser == 'b'
    lone = [p for p in st['b'] if p and p.upper() != 'K' and _white(p) == opp_white]
    if not lone:
        finish_game(g, '1/2-1/2', 'timeout-draw')
    elif loser == 'w':
        finish_game(g, '0-1', 'timeout', g['black'])
    else:
        finish_game(g, '1-0', 'timeout', g['white'])
    return g, None


# ---------------- Giải đấu: loại trực tiếp ----------------

def bracket_start(players):
    """players: list key theo thứ tự đăng ký. Trả matches vòng 1 (kèm bye tự thắng)."""
    matches = []
    n = len(players)
    i = 0
    slot = 0
    while i < n:
        if i + 1 < n:
            matches.append({'id': 'm1-%d' % slot, 'round': 1, 'p1': players[i], 'p2': players[i + 1],
                            'game_id': None, 'winner': None, 'status': 'pending'})
            i += 2
        else:
            matches.append({'id': 'm1-%d' % slot, 'round': 1, 'p1': players[i], 'p2': None,
                            'game_id': None, 'winner': players[i], 'status': 'bye'})
            i += 1
        slot += 1
    return matches


def bracket_advance(tour, match, winner_key):
    """Sau khi 1 trận xong: ghi người thắng, ghép vòng sau, trả game mới cần tạo (hoặc None)."""
    match['winner'] = winner_key
    match['status'] = 'done'
    rnd = match['round']
    same = [m for m in tour['matches'] if m['round'] == rnd and m['status'] in ('done', 'bye')]
    # Ghép cặp vòng sau theo thứ tự: 2 trận xong liên tiếp tạo 1 trận vòng sau
    done_ids = [m['id'] for m in tour['matches'] if m['round'] == rnd and m['status'] == 'done']
    if len(done_ids) % 2 == 1:
        return None  # chờ trận còn lại của cặp
    # Tìm 2 trận vừa xong chưa được ghép
    paired = set()
    for m in tour['matches']:
        if m['round'] == rnd + 1:
            paired.add(m.get('from1'))
            paired.add(m.get('from2'))
    ready = [m for m in tour['matches'] if m['round'] == rnd and m['status'] == 'done'
             and m['id'] not in paired]
    if len(ready) < 2:
        return None
    a, b = ready[0], ready[1]
    nxt = {'id': 'm%d-%d' % (rnd + 1, len([m for m in tour['matches'] if m['round'] == rnd + 1])),
           'round': rnd + 1, 'p1': a['winner'], 'p2': b['winner'], 'game_id': None,
           'winner': None, 'status': 'pending', 'from1': a['id'], 'from2': b['id']}
    tour['matches'].append(nxt)
    return nxt


def bracket_champion(tour):
    if not tour['matches']:
        return None
    last_round = max(m['round'] for m in tour['matches'])
    finals = [m for m in tour['matches'] if m['round'] == last_round]
    if len(finals) == 1 and finals[0]['status'] == 'done':
        return finals[0]['winner']
    return None
