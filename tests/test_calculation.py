import pytest
import pandas as pd
import numpy as np
from utils.calculation import (
    calculate_relative_score,
    calculate_absolute_score,
    calculate_org_absolute_score,
    calculate_relative_absolute_score,
    calculate_scores
)

def test_calculate_relative_score():
    # 正常系: 3人中1位が100点、2位が66.6点、3位が33.3点
    weight = 100
    all_values = [10, 20, 30]
    assert calculate_relative_score(30, all_values, weight) == 100
    assert calculate_relative_score(20, all_values, weight) == pytest.approx(66.6666, 0.001)
    assert calculate_relative_score(10, all_values, weight) == pytest.approx(33.3333, 0.001)
    
    # 異常系: 実績0は0点
    assert calculate_relative_score(0, all_values, weight) == 0
    assert calculate_relative_score(-5, all_values, weight) == 0
    assert calculate_relative_score(None, all_values, weight) == 0

def test_calculate_absolute_score():
    weight = 100
    target = 10
    # 正常系
    assert calculate_absolute_score(5, target, weight) == 50
    assert calculate_absolute_score(10, target, weight) == 100
    # キャップ: 目標以上でも配点まで
    assert calculate_absolute_score(20, target, weight) == 100
    # 異常系: 目標0や負数は0点
    assert calculate_absolute_score(5, 0, weight) == 0

def test_calculate_org_absolute_score():
    weight = 100
    target = 100
    # All or Nothing
    assert calculate_org_absolute_score(100, target, weight) == 100
    assert calculate_org_absolute_score(110, target, weight) == 100
    assert calculate_org_absolute_score(99.9, target, weight) == 0

def test_calculate_relative_absolute_score():
    weight = 100
    all_values = [10, 20, 30] # 合計60
    all_targets = [20, 20, 20] # 合計60 -> 全体平均率 1.0
    # 最低保証ライン: 1.0 * 0.5 = 0.5 (50%)
    
    # A: 30/20 = 1.5 (1位) -> 100点
    assert calculate_relative_absolute_score(30, 20, all_values, all_targets, weight) == 100
    # B: 20/20 = 1.0 (2位) -> ランキングは 100 - (2-1)*33.3 = 66.6
    assert calculate_relative_absolute_score(20, 20, all_values, all_targets, weight) == pytest.approx(66.6666, 0.001)
    # C: 5/20 = 0.25 (3位) -> 達成率0.25 < 最低保証0.5 -> ランキング値 (33.3)
    assert calculate_relative_absolute_score(5, 20, [10, 20, 5], all_targets, weight) == pytest.approx(33.3333, 0.001)
    # D: 11/20 = 0.55 (3位) -> 達成率0.55 > 最低保証0.5 -> 最低保証適用 (50点)
    # 3人ランキングの3位は33.3点だが、保証で50点になる
    assert calculate_relative_absolute_score(11, 20, [20, 30, 11], [20, 20, 20], weight) == 50

def test_independent_achievement_logic():
    # 達成評価（独立加算）の統合テスト
    config = {
        'items': [{
            'name': 'テスト項目',
            'weight': 10,
            'evaluation_type': 'independent_achievement',
            'targets': {'personal_target': 10, 'team_target': 100},
            'weights': {'w1': 10, 'w2': 20, 'w3': 30},
            'shop_targets': {'店舗A': 500}
        }]
    }
    
    from utils.calculation import normalize_col_name
    
    df = pd.DataFrame({
        normalize_col_name('スタッフ名'): [normalize_col_name('スタッフ1'), normalize_col_name('スタッフ2')],
        normalize_col_name('店舗名'): [normalize_col_name('店舗A'), normalize_col_name('店舗A')],
        normalize_col_name('チーム名'): [normalize_col_name('チームX'), normalize_col_name('チームX')],
        normalize_col_name('テスト項目'): [12, 8] # 合計20
    })
    
    config['items'][0]['targets']['team_target'] = 15
    
    # スタッフ1: 個人(12>=10)達成(+10) & チーム(20>=15)達成(+20) -> 30点
    # スタッフ2: 個人未達成 & チーム達成(+20) -> 20点
    res = calculate_scores(df, config)
    assert res.loc[0, 'テスト項目_score'] == 30
    assert res.loc[1, 'テスト項目_score'] == 20
    
    # 4. 店舗判定のテスト
    # 店舗Aの目標 500. 20 < 500 なので未達成。
    # 目標を 15 に下げて達成させる
    config['items'][0]['shop_targets']['店舗A'] = 15
    # スタッフ1: 個人(12>=10)+10, チーム(20>=15)+20, 店舗(20>=15)+30 -> 60点
    res = calculate_scores(df, config)
    assert res.loc[0, 'テスト項目_score'] == 60
