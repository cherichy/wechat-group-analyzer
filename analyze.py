import csv
import sys
import io
from collections import defaultdict
from datetime import datetime

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

filename = r'chat_export_CuTeDSL & cuTile 交流群_57327409534@chatroom.csv'

msg_count = defaultdict(int)
last_msg_time = {}
first_msg_time = {}
user_names = {}  # id -> nickname (latest)

total_lines = 0
min_time = None
max_time = None

with open(filename, 'r', encoding='utf-8-sig', errors='replace') as f:
    reader = csv.reader(f)
    header = next(reader)
    for row in reader:
        if len(row) < 5:
            continue
        total_lines += 1
        time_str = row[0].strip()
        nickname = row[1].strip()
        sender_id = row[2].strip()

        try:
            t = datetime.strptime(time_str, '%Y-%m-%d %H:%M:%S')
        except:
            continue

        if min_time is None or t < min_time:
            min_time = t
        if max_time is None or t > max_time:
            max_time = t

        msg_count[sender_id] += 1
        user_names[sender_id] = nickname

        if sender_id not in first_msg_time or t < first_msg_time[sender_id]:
            first_msg_time[sender_id] = t
        if sender_id not in last_msg_time or t > last_msg_time[sender_id]:
            last_msg_time[sender_id] = t

print("=" * 60)
print("基本统计")
print("=" * 60)
print(f"时间范围: {min_time} ~ {max_time}")
days_span = (max_time - min_time).days
print(f"跨度: {days_span} 天 (~{days_span/30:.1f} 个月)")
print(f"总消息数: {total_lines}")
print(f"发言人数(唯一): {len(msg_count)}")
print(f"群总人数: ~500 (用户提供)")
print(f"从未发言人数: ~{500 - len(msg_count)}")
print()

# Sort by message count descending
sorted_users = sorted(msg_count.items(), key=lambda x: x[1], reverse=True)

print("=" * 60)
print("发言量 Top 30")
print("=" * 60)
for i, (uid, count) in enumerate(sorted_users[:30]):
    name = user_names[uid]
    last = last_msg_time[uid].strftime('%Y-%m-%d')
    first = first_msg_time[uid].strftime('%Y-%m-%d')
    print(f"{i+1:3}. {name:20s} | {count:5d}条 | 首次: {first} | 最后: {last}")

print()

# Distribution
print("=" * 60)
print("发言量分布")
print("=" * 60)
brackets_labels = [
    (1, 1, "1条"),
    (2, 2, "2条"),
    (3, 4, "3-4条"),
    (5, 9, "5-9条"),
    (10, 19, "10-19条"),
    (20, 49, "20-49条"),
    (50, 99, "50-99条"),
    (100, 199, "100-199条"),
    (200, 499, "200-499条"),
    (500, 999, "500-999条"),
    (1000, 999999, "1000+条"),
]
for lo, hi, label in brackets_labels:
    n = sum(1 for uid, c in msg_count.items() if lo <= c <= hi)
    bar = "#" * n
    print(f"  {label:10s}: {n:4d}人  {bar}")

print()

# Last message time distribution
now = max_time
print("=" * 60)
print(f"最后发言距今分布 (基准: {now.strftime('%Y-%m-%d')})")
print("=" * 60)
time_brackets = [
    (0, 7, "7天内"),
    (8, 30, "8-30天"),
    (31, 60, "1-2月"),
    (61, 90, "2-3月"),
    (91, 180, "3-6月"),
    (181, 365, "6-12月"),
    (366, 99999, "超过1年"),
]
cumulative = 0
for lo, hi, label in time_brackets:
    n = sum(1 for uid in msg_count if lo <= (now - last_msg_time[uid]).days <= hi)
    cumulative += n
    print(f"  {label:10s}: {n:4d}人 (累计: {cumulative})")

print()

# Monthly activity trend
print("=" * 60)
print("月度活跃人数趋势")
print("=" * 60)
monthly_users = defaultdict(set)
monthly_msgs = defaultdict(int)
for uid in msg_count:
    # We need to re-scan... let's just compute from stored data
    pass

# Re-scan for monthly data
with open(filename, 'r', encoding='utf-8-sig', errors='replace') as f:
    reader = csv.reader(f)
    next(reader)
    for row in reader:
        if len(row) < 5:
            continue
        time_str = row[0].strip()
        sender_id = row[2].strip()
        try:
            t = datetime.strptime(time_str, '%Y-%m-%d %H:%M:%S')
        except:
            continue
        month_key = t.strftime('%Y-%m')
        monthly_users[month_key].add(sender_id)
        monthly_msgs[month_key] += 1

for month in sorted(monthly_users.keys()):
    n_users = len(monthly_users[month])
    n_msgs = monthly_msgs[month]
    bar = "#" * (n_users // 2)
    print(f"  {month}: {n_users:4d}人活跃, {n_msgs:5d}条消息  {bar}")

print()

# Identify candidates for removal
# Strategy: people who sent very few messages AND haven't spoken recently
print("=" * 60)
print("清理建议分析")
print("=" * 60)

# Score each user: lower = more likely to be removed
# Score = msg_count_weight + recency_weight
user_scores = {}
for uid in msg_count:
    days_since_last = (now - last_msg_time[uid]).days
    count = msg_count[uid]
    # Simple scoring: log of msg count + recency bonus
    import math
    score = math.log2(count + 1) * 10 + max(0, (365 - days_since_last)) / 365 * 50
    user_scores[uid] = score

sorted_by_score = sorted(user_scores.items(), key=lambda x: x[1])

# First: people who never spoke (~500-unique_count) are auto-remove candidates
never_spoke = 500 - len(msg_count)
need_to_remove = 500 - 350  # 150 people
from_active = max(0, need_to_remove - never_spoke)

print(f"目标: 500 -> 350 (需移除 ~{need_to_remove} 人)")
print(f"从未发言: ~{never_spoke} 人 (优先移除)")
print(f"还需从发言者中移除: ~{from_active} 人")
print()

if from_active > 0:
    print(f"建议移除的低活跃发言者 (活跃度最低的 {from_active} 人):")
    print("-" * 60)
    for i, (uid, score) in enumerate(sorted_by_score[:from_active]):
        name = user_names[uid]
        count = msg_count[uid]
        last = last_msg_time[uid].strftime('%Y-%m-%d')
        days_ago = (now - last_msg_time[uid]).days
        print(f"  {i+1:3}. {name:20s} | {count:3d}条 | 最后发言: {last} ({days_ago}天前) | 得分: {score:.1f}")

print()
print("=" * 60)
print("得分说明: 综合消息数量(log2) + 最近活跃度，得分越低越不活跃")
print("=" * 60)

# Export per-user stats to CSV
import math
output_csv = 'member_stats.csv'
with open(output_csv, 'w', encoding='utf-8-sig', newline='') as f:
    writer = csv.writer(f)
    writer.writerow(['昵称', '微信ID', '发言总数', '首次发言', '最后发言', '距今天数', '活跃得分'])
    for uid, score in sorted(user_scores.items(), key=lambda x: x[1], reverse=True):
        name = user_names[uid]
        count = msg_count[uid]
        first = first_msg_time[uid].strftime('%Y-%m-%d')
        last = last_msg_time[uid].strftime('%Y-%m-%d')
        days_ago = (now - last_msg_time[uid]).days
        writer.writerow([name, uid, count, first, last, days_ago, round(score, 1)])
print(f"\n已导出 {len(user_scores)} 条记录到 {output_csv}")
