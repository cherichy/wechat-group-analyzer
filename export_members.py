import csv
import json
import sys
import io
import urllib.request

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

BASE = 'http://127.0.0.1:5200/api/v1'

# Load group members
with open('chatroom_info_raw.json', 'rb') as f:
    data = json.loads(f.read().decode('utf-8'))

members = data['data']['users']
owner_id = data['data']['owner']
print(f'群成员: {len(members)} 人')

# Source 1: member_activity (names of people who sent messages)
print('正在获取 member_activity...')
activity_names = {}
url = f'{BASE}/analysis/member_activity/57327409534@chatroom'
try:
    with urllib.request.urlopen(url, timeout=30) as resp:
        adata = json.loads(resp.read().decode('utf-8'))
    items = adata.get('data', adata) if isinstance(adata, dict) else adata
    if isinstance(items, list):
        for item in items:
            pid = item.get('platformId', '')
            name = item.get('name', '')
            if pid and pid != 'self' and name:
                activity_names[pid] = name
    print(f'  member_activity: {len(activity_names)} 个发言者昵称')
except Exception as e:
    print(f'  member_activity 失败: {e}')

# Source 2: contacts/:id for each member (get nickName, alias)
print(f'正在逐个查询 {len(members)} 个成员的联系人信息...')
contacts_map = {}
for i, m in enumerate(members):
    uid = m['userName']
    url = f'{BASE}/contacts/{uid}'
    try:
        with urllib.request.urlopen(url, timeout=10) as resp:
            raw = resp.read().decode('utf-8')
        cdata = json.loads(raw)
        c = cdata.get('data', cdata) if isinstance(cdata, dict) else {}
        contacts_map[uid] = {
            'nickName': c.get('nickName', ''),
            'alias': c.get('alias', ''),
            'remark': c.get('remark', ''),
        }
    except:
        pass
    if (i + 1) % 50 == 0:
        print(f'  已查询 {i+1}/{len(members)}')

print(f'联系人查询完成: {len(contacts_map)} 人有数据')

# Export CSV
output = 'group_members_full.csv'
with open(output, 'w', encoding='utf-8-sig', newline='') as f:
    writer = csv.writer(f)
    writer.writerow(['序号', '微信ID', '微信昵称', '群昵称', '微信号(alias)', '备注', '活跃昵称(来自聊天)', '是否群主'])
    for i, m in enumerate(members, 1):
        uid = m['userName']
        display = m['displayName']
        info = contacts_map.get(uid, {})
        nick = info.get('nickName', '')
        alias = info.get('alias', '')
        remark = info.get('remark', '')
        chat_name = activity_names.get(uid, '')
        is_owner = '是' if uid == owner_id else ''
        writer.writerow([i, uid, nick, display, alias, remark, chat_name, is_owner])

# Stats
has_nick = sum(1 for m in members if contacts_map.get(m['userName'], {}).get('nickName', ''))
has_display = sum(1 for m in members if m['displayName'])
has_chat = sum(1 for m in members if m['userName'] in activity_names)
has_any = sum(1 for m in members if (
    contacts_map.get(m['userName'], {}).get('nickName', '') or
    m['displayName'] or
    m['userName'] in activity_names
))
no_name = len(members) - has_any

print(f'\n=== 导出统计 ===')
print(f'有微信昵称 (contacts API): {has_nick}')
print(f'有群昵称 (chatroom API): {has_display}')
print(f'有活跃昵称 (member_activity): {has_chat}')
print(f'至少有一个名字: {has_any}')
print(f'完全无名 (只有wxid): {no_name}')
print(f'已导出到 {output}')
