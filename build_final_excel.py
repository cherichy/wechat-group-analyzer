"""
Build one comprehensive Excel file with ALL 499 group members.
Data sources:
  1. chatroom_info_raw.json  - member list (userName, displayName), owner
  2. contacts API cache      - nickName, alias, remark (from group_members_full.csv)
  3. member_stats.csv        - chat activity stats (message count, first/last message, activity score)
  4. messages JSON           - avatar URLs (smallHeadURL, bigHeadURL) per sender
  5. member_activity         - active nicknames (from group_members_full.csv)
"""

import csv
import json
import sys
import io
import os

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# ---- 1. Load group member list ----
print('Loading chatroom_info_raw.json...')
with open('chatroom_info_raw.json', 'rb') as f:
    chatroom = json.loads(f.read().decode('utf-8'))
members = chatroom['data']['users']
owner_id = chatroom['data']['owner']
print(f'  {len(members)} members, owner={owner_id}')

# ---- 2. Load contacts & activity names from group_members_full.csv ----
print('Loading group_members_full.csv...')
contacts_map = {}  # uid -> {nickName, displayName, alias, remark, chatName, isOwner}
with open('group_members_full.csv', 'r', encoding='utf-8-sig') as f:
    reader = csv.DictReader(f)
    for row in reader:
        uid = row['微信ID']
        contacts_map[uid] = {
            'nickName': row.get('微信昵称', ''),
            'displayName': row.get('群昵称', ''),
            'alias': row.get('微信号(alias)', ''),
            'remark': row.get('备注', ''),
            'chatName': row.get('活跃昵称(来自聊天)', ''),
            'isOwner': row.get('是否群主', ''),
        }
print(f'  {len(contacts_map)} entries loaded')

# ---- 3. Load chat stats from member_stats.csv ----
print('Loading member_stats.csv...')
stats_map = {}  # uid -> {msgCount, firstMsg, lastMsg, daysSince, score}
with open('member_stats.csv', 'r', encoding='utf-8-sig') as f:
    reader = csv.DictReader(f)
    for row in reader:
        uid = row.get('微信ID', '').strip()
        if not uid:
            continue
        stats_map[uid] = {
            'msgCount': int(row.get('发言总数', 0)),
            'firstMsg': row.get('首次发言', ''),
            'lastMsg': row.get('最后发言', ''),
            'daysSince': row.get('距今天数', ''),
            'score': row.get('活跃得分', ''),
        }
print(f'  {len(stats_map)} speakers with stats')

# ---- 4. Extract avatar URLs from messages JSON ----
print('Loading messages JSON for avatar URLs (this may take a moment)...')
avatar_map = {}  # sender_id -> {smallHeadURL, bigHeadURL, senderName}

json_file = 'messages_57327409534@chatroom (1).json'
if os.path.exists(json_file):
    with open(json_file, 'rb') as f:
        raw = f.read()
    msgdata = json.loads(raw.decode('utf-8'))
    messages = msgdata.get('data', [])
    print(f'  {len(messages)} messages in JSON')
    
    for msg in messages:
        sender = msg.get('sender', '')
        if not sender or sender == '57327409534@chatroom':
            continue
        # Only update if we get a non-empty URL (prefer later/newer entries)
        small = msg.get('smallHeadURL', '')
        big = msg.get('bigHeadURL', '')
        name = msg.get('senderName', '')
        
        if sender not in avatar_map:
            avatar_map[sender] = {'smallHeadURL': '', 'bigHeadURL': '', 'senderName': ''}
        
        if small:
            avatar_map[sender]['smallHeadURL'] = small
        if big:
            avatar_map[sender]['bigHeadURL'] = big
        if name:
            avatar_map[sender]['senderName'] = name
    
    print(f'  {len(avatar_map)} unique senders with avatar data')
    has_small = sum(1 for v in avatar_map.values() if v['smallHeadURL'])
    has_big = sum(1 for v in avatar_map.values() if v['bigHeadURL'])
    print(f'  {has_small} with smallHeadURL, {has_big} with bigHeadURL')
else:
    print(f'  WARNING: {json_file} not found!')

# ---- 5. Build comprehensive data for all 499 members ----
print('\nBuilding comprehensive member list...')

rows = []
for i, m in enumerate(members, 1):
    uid = m['userName']
    display_from_api = m.get('displayName', '')
    
    c = contacts_map.get(uid, {})
    s = stats_map.get(uid, {})
    a = avatar_map.get(uid, {})
    
    nick = c.get('nickName', '')
    display = c.get('displayName', '') or display_from_api
    alias = c.get('alias', '')
    remark = c.get('remark', '')
    chat_name = c.get('chatName', '')
    is_owner = '是' if uid == owner_id else ''
    
    # Best recognizable name: prefer remark > chatName > nickName > displayName > uid
    best_name = remark or chat_name or nick or display or uid
    
    msg_count = s.get('msgCount', 0)
    first_msg = s.get('firstMsg', '')
    last_msg = s.get('lastMsg', '')
    days_since = s.get('daysSince', '')
    score = s.get('score', '')
    
    small_url = a.get('smallHeadURL', '')
    big_url = a.get('bigHeadURL', '')
    
    # Activity status
    if msg_count > 0:
        activity = '活跃'
    else:
        activity = '从未发言'
    
    # Cleanup recommendation
    if is_owner:
        recommend = '保留(群主)'
    elif msg_count > 0:
        recommend = '保留'
    else:
        recommend = '建议移除'
    
    rows.append({
        '序号': i,
        '微信ID': uid,
        '可辨识昵称': best_name,
        '微信昵称': nick,
        '群昵称': display,
        '活跃昵称': chat_name,
        '微信号': alias,
        '备注': remark,
        '头像URL(小)': small_url,
        '头像URL(大)': big_url,
        '发言总数': msg_count,
        '首次发言': first_msg,
        '最后发言': last_msg,
        '距今天数': days_since,
        '活跃得分': score,
        '活跃状态': activity,
        '清理建议': recommend,
        '是否群主': is_owner,
    })

# ---- 6. Export to CSV (Excel-compatible) ----
output = 'group_members_comprehensive.csv'
fieldnames = [
    '序号', '微信ID', '可辨识昵称', '微信昵称', '群昵称', '活跃昵称',
    '微信号', '备注', '头像URL(小)', '头像URL(大)',
    '发言总数', '首次发言', '最后发言', '距今天数', '活跃得分',
    '活跃状态', '清理建议', '是否群主',
]

with open(output, 'w', encoding='utf-8-sig', newline='') as f:
    writer = csv.DictWriter(f, fieldnames=fieldnames)
    writer.writeheader()
    # Sort: owner first, then by message count descending, then never-spoke at bottom
    rows.sort(key=lambda r: (
        0 if r['是否群主'] == '是' else 1,
        0 if r['发言总数'] > 0 else 1,
        -r['发言总数'],
    ))
    # Re-number after sort
    for idx, r in enumerate(rows, 1):
        r['序号'] = idx
        writer.writerow(r)

# ---- 7. Summary stats ----
total = len(rows)
active = sum(1 for r in rows if r['活跃状态'] == '活跃')
inactive = total - active
has_avatar = sum(1 for r in rows if r['头像URL(小)'])
has_nick = sum(1 for r in rows if r['微信昵称'])
has_display = sum(1 for r in rows if r['群昵称'])
has_chat_name = sum(1 for r in rows if r['活跃昵称'])
has_alias = sum(1 for r in rows if r['微信号'])
remove = sum(1 for r in rows if r['清理建议'] == '建议移除')
keep = total - remove

print(f'\n{"="*50}')
print(f'导出完成: {output}')
print(f'{"="*50}')
print(f'总人数: {total}')
print(f'活跃(有发言): {active}')
print(f'从未发言: {inactive}')
print(f'建议移除: {remove}')
print(f'建议保留: {keep}')
print(f'')
print(f'数据完整度:')
print(f'  有微信昵称: {has_nick}/{total}')
print(f'  有群昵称: {has_display}/{total}')
print(f'  有活跃昵称: {has_chat_name}/{total}')
print(f'  有微信号: {has_alias}/{total}')
print(f'  有头像URL: {has_avatar}/{total}')
print(f'')
print(f'排序: 群主 > 活跃成员(按发言数降序) > 从未发言成员')
