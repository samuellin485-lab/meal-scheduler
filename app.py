# 套用精準次數控制與分數偏移
for p, adj in custom_adjustments.items():
    if p in members:
        b_adj = adj.get("飯前", 0)
        a_adj = adj.get("飯後", 0)
        
        # 1. 飯前次數控制：若設定 -1，強迫該成員飯前總次數不得超過 (預設平均 + b_adj)
        if b_adj != 0:
            target_b = max(0, math.floor(avg_before + b_adj))
            model.Add(sum(x[p, d, "飯前"] for d in dates) <= target_b)
            
        # 2. 飯後次數控制：若設定 -1，強迫該成員飯後總次數不得超過 (預設平均 + a_adj)
        if a_adj != 0:
            target_a = max(0, math.floor(avg_after + a_adj))
            model.Add(sum(x[p, d, "飯後"] for d in dates) <= target_a)

        # 3. 防止演算法「拿飯前補飯後」：
        # 如果設定少排飯後(a_adj < 0)，同時限制他的飯前次數不得超過全體平均，防止系統自動多排飯前給他
        if a_adj < 0 and b_adj == 0:
            model.Add(sum(x[p, d, "飯前"] for d in dates) <= math.floor(avg_before))
