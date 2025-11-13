import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import pulp
from typing import List, Dict, Tuple
import warnings

warnings.filterwarnings('ignore')

plt.rcParams['font.sans-serif'] = ['SimHei']  # 支持中文显示
plt.rcParams['axes.unicode_minus'] = False


class WildfireEquipmentModel:
    def __init__(self):
        # 基础参数配置 - 使用更合理的数值
        self.params = {
            # 区域参数
            'R': 5,  # 区域数量
            'lambda_i': [8, 12, 6, 10, 7],  # 各区域年平均火灾次数

            # 火灾规模参数
            'A_mean': 1.5,  # 平均火场面积(km²)
            'A_std': 1.0,  # 面积标准差
            'alpha1': 2.0,  # 基础小队数
            'alpha2': 0.3,  # 每平方公里增加的小队数

            # 设备参数
            'S_SSA': 0.8,  # 单架SSA覆盖面积(km²)
            'T_cycle': 1.5,  # 巡航周期(小时)
            'R_max': 0.3,  # 最大重访时间(小时)
            'R_cov': 15,  # 中继覆盖半径(km)
            'beta': 1.3,  # 中继冗余系数

            # 成本参数
            'p_SSA': 80000,  # SSA单价($)
            'p_R': 25000,  # 中继单价($)

            # 安全系数
            'gamma': 1.2  # 安全冗余系数
        }

    def generate_fire_events(self, year=0, growth_rate=0.03):
        """生成一年内的火灾事件"""
        print(f"生成第{year}年火灾事件...")
        events = []
        lambda_i = [lam * (1 + growth_rate) ** year for lam in self.params['lambda_i']]

        for region in range(self.params['R']):
            # 生成火灾次数 - 泊松分布
            n_fires = np.random.poisson(lambda_i[region])

            for _ in range(n_fires):
                # 随机起火时间 (0-8759小时)
                start_time = np.random.randint(0, 8760)

                # 火场面积 - 对数正态分布，避免极端值
                A_e = np.random.lognormal(
                    mean=np.log(self.params['A_mean']),
                    sigma=np.log(self.params['A_std'])
                )
                A_e = max(0.1, min(10.0, A_e))  # 限制在合理范围

                # 计算相关参数
                K_e = max(1, self.params['alpha1'] + self.params['alpha2'] * A_e)
                T_e = max(4, min(48, A_e * 8))  # 持续时间4-48小时
                L_e = 3 * np.sqrt(A_e * np.pi)  # 火线长度

                events.append({
                    'region': region,
                    'start_time': start_time,
                    'duration': T_e,
                    'area': round(A_e, 2),
                    'squads': round(K_e, 1),
                    'fireline_length': round(L_e, 2)
                })

        print(f"共生成{len(events)}次火灾事件")
        return events

    def calculate_equipment_demand(self, fire_events):
        """计算单次火灾的设备需求"""
        print("计算设备需求...")
        for event in fire_events:
            A_e = event['area']
            K_e = event['squads']

            # SSA需求计算
            n_SSA1 = np.ceil(A_e / self.params['S_SSA'] *
                             self.params['T_cycle'] / self.params['R_max'])
            n_SSA2 = np.ceil(K_e / 2)  # 假设每架服务2个小队
            event['n_SSA'] = int(max(n_SSA1, n_SSA2))

            # 中继需求计算
            L_e = event['fireline_length']
            n_R = np.ceil(self.params['beta'] * L_e / (2 * self.params['R_cov']))
            event['n_R'] = int(max(1, n_R))  # 至少1架

        return fire_events

    def create_demand_timeseries(self, fire_events):
        """创建年度时间序列需求"""
        print("创建时间序列需求...")
        hours_per_year = 8760
        D_SSA = np.zeros(hours_per_year)
        D_R = np.zeros(hours_per_year)

        for event in fire_events:
            start = event['start_time']
            duration = event['duration']
            end = min(hours_per_year, start + duration)

            # 确保时间范围有效
            if start < hours_per_year:
                for t in range(int(start), int(end)):
                    if t < hours_per_year:
                        D_SSA[t] += event['n_SSA']
                        D_R[t] += event['n_R']

        return D_SSA, D_R

    def optimize_equipment_config(self, D_SSA_max, D_R_max):
        """优化设备配置"""
        print("优化设备配置...")

        # 计算考虑安全系数的需求
        demand_SSA = np.ceil(self.params['gamma'] * D_SSA_max)
        demand_R = np.ceil(self.params['gamma'] * D_R_max)

        # 创建优化问题
        prob = pulp.LpProblem("Equipment_Optimization", pulp.LpMinimize)

        # 决策变量
        x_SSA = pulp.LpVariable("x_SSA", lowBound=demand_SSA, cat='Integer')
        x_R = pulp.LpVariable("x_R", lowBound=demand_R, cat='Integer')

        # 目标函数：最小化总成本
        prob += self.params['p_SSA'] * x_SSA + self.params['p_R'] * x_R

        # 求解
        prob.solve()

        result = {
            'x_SSA_opt': int(x_SSA.varValue),
            'x_R_opt': int(x_R.varValue),
            'total_cost': pulp.value(prob.objective),
            'peak_demand_SSA': D_SSA_max,
            'peak_demand_R': D_R_max,
            'status': pulp.LpStatus[prob.status]
        }

        return result

    def multi_year_expansion(self, base_config, years=10, growth_rate=0.03):
        """多年扩容规划"""
        print("进行多年扩容分析...")
        results = []

        # 初始存量
        F_SSA = base_config['x_SSA_opt']
        F_R = base_config['x_R_opt']

        for year in range(years + 1):
            # 生成该年火灾事件
            events = self.generate_fire_events(year, growth_rate)
            events_with_demand = self.calculate_equipment_demand(events)
            D_SSA, D_R = self.create_demand_timeseries(events_with_demand)

            # 年度高峰需求
            D_SSA_max = np.max(D_SSA) if len(D_SSA) > 0 else 0
            D_R_max = np.max(D_R) if len(D_R) > 0 else 0

            # 考虑安全系数的需求
            required_SSA = np.ceil(self.params['gamma'] * D_SSA_max)
            required_R = np.ceil(self.params['gamma'] * D_R_max)

            # 采购决策
            buy_SSA = max(0, required_SSA - F_SSA)
            buy_R = max(0, required_R - F_R)

            # 更新存量
            F_SSA += buy_SSA
            F_R += buy_R

            # 年度成本
            annual_cost = (buy_SSA * self.params['p_SSA'] +
                           buy_R * self.params['p_R'])

            results.append({
                'year': year,
                'peak_demand_SSA': D_SSA_max,
                'peak_demand_R': D_R_max,
                'required_SSA': required_SSA,
                'required_R': required_R,
                'buy_SSA': int(buy_SSA),
                'buy_R': int(buy_R),
                'inventory_SSA': int(F_SSA),
                'inventory_R': int(F_R),
                'annual_cost': annual_cost
            })

        return pd.DataFrame(results)


class RelayDeploymentOptimizer:
    """中继无人机布设优化"""

    def __init__(self):
        self.R0 = 20  # 基础通信半径(km)

    def generate_scenario(self):
        """生成示例场景"""
        # EOC位置
        eoc = (0, 0)

        # 前线小队位置 - 在50x50km区域内
        frontlines = [
            (15, 10), (25, 15), (35, 8),
            (20, 25), (30, 30), (40, 20)
        ]

        # 候选中继点 - 网格分布
        candidates = []
        for x in range(5, 45, 10):
            for y in range(5, 35, 10):
                candidates.append((x, y))

        return eoc, frontlines, candidates

    def calculate_distance(self, pos1, pos2):
        """计算两点间距离"""
        return np.sqrt((pos1[0] - pos2[0]) ** 2 + (pos1[1] - pos2[1]) ** 2)

    def optimize_relay_deployment(self, eoc, frontlines, candidates):
        """优化中继部署"""
        print("优化中继部署...")

        prob = pulp.LpProblem("Relay_Deployment", pulp.LpMinimize)

        # 决策变量
        y = {j: pulp.LpVariable(f"y_{j}", cat='Binary')
             for j in range(len(candidates))}

        # 目标函数：最小化中继数量
        prob += pulp.lpSum([y[j] for j in range(len(candidates))])

        # 约束条件：每个前线小队必须被覆盖
        for k, frontline in enumerate(frontlines):
            # 检查是否能直接连接到EOC
            direct_connect = 1 if self.calculate_distance(frontline, eoc) <= self.R0 else 0

            # 找到能覆盖该小队的中继候选点
            covering_relays = []
            for j, candidate in enumerate(candidates):
                if self.calculate_distance(frontline, candidate) <= self.R0:
                    covering_relays.append(y[j])

            # 约束：要么直连EOC，要么至少有一个中继覆盖
            prob += (pulp.lpSum(covering_relays) + direct_connect >= 1)

        # 求解
        prob.solve()

        # 提取结果
        deployed_relays = []
        for j in range(len(candidates)):
            if pulp.value(y[j]) == 1:
                deployed_relays.append(candidates[j])

        return {
            'deployed_positions': deployed_relays,
            'num_relays': len(deployed_relays),
            'status': pulp.LpStatus[prob.status]
        }


class Visualization:
    """结果可视化"""

    @staticmethod
    def plot_demand_timeseries(D_SSA, D_R):
        """绘制时间序列需求"""
        fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(12, 8))

        # 只显示前1000小时以便观察
        display_hours = min(1000, len(D_SSA))

        # SSA需求
        ax1.plot(range(display_hours), D_SSA[:display_hours], alpha=0.7, linewidth=1)
        ax1.axhline(y=np.max(D_SSA), color='r', linestyle='--',
                    label=f'峰值需求: {np.max(D_SSA):.1f}')
        ax1.set_ylabel('SSA无人机需求')
        ax1.set_title('SSA无人机需求时间序列')
        ax1.legend()
        ax1.grid(True, alpha=0.3)

        # 中继需求
        ax2.plot(range(display_hours), D_R[:display_hours], alpha=0.7, linewidth=1)
        ax2.axhline(y=np.max(D_R), color='r', linestyle='--',
                    label=f'峰值需求: {np.max(D_R):.1f}')
        ax2.set_ylabel('中继无人机需求')
        ax2.set_xlabel('年度时间 (小时)')
        ax2.set_title('中继无人机需求时间序列')
        ax2.legend()
        ax2.grid(True, alpha=0.3)

        plt.tight_layout()
        plt.show()

    @staticmethod
    def plot_multi_year_analysis(multi_year_df):
        """绘制多年分析结果"""
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(15, 6))

        years = multi_year_df['year']

        # 设备存量趋势
        ax1.plot(years, multi_year_df['inventory_SSA'], 'o-',
                 label='SSA库存', linewidth=2, markersize=6)
        ax1.plot(years, multi_year_df['inventory_R'], 's-',
                 label='中继库存', linewidth=2, markersize=6)
        ax1.set_xlabel('年份')
        ax1.set_ylabel('无人机数量')
        ax1.set_title('设备库存年度变化')
        ax1.legend()
        ax1.grid(True, alpha=0.3)

        # 成本趋势
        ax2.bar(years, multi_year_df['annual_cost'], alpha=0.7,
                label='年度成本', color='skyblue')
        ax2.plot(years, multi_year_df['annual_cost'].cumsum(), 'ro-',
                 label='累计成本', linewidth=2, markersize=6)
        ax2.set_xlabel('年份')
        ax2.set_ylabel('成本 ($)')
        ax2.set_title('成本分析')
        ax2.legend()
        ax2.grid(True, alpha=0.3)

        # 在柱状图上标注数值
        for i, cost in enumerate(multi_year_df['annual_cost']):
            if cost > 0:
                ax2.text(i, cost + 10000, f'${cost / 1000:.0f}K',
                         ha='center', va='bottom', fontsize=8)

        plt.tight_layout()
        plt.show()

    @staticmethod
    def plot_relay_deployment(eoc, frontlines, candidates, deployed_relays):
        """绘制中继部署方案"""
        fig, ax = plt.subplots(figsize=(10, 8))

        # 绘制EOC
        ax.scatter(*eoc, s=200, c='red', marker='s', label='指挥中心', zorder=5)
        ax.text(eoc[0], eoc[1] + 2, 'EOC', ha='center', fontweight='bold')

        # 绘制前线小队
        frontline_x, frontline_y = zip(*frontlines)
        ax.scatter(frontline_x, frontline_y, s=100, c='blue',
                   marker='^', label='前线小队', zorder=4)

        # 绘制候选点
        candidate_x, candidate_y = zip(*candidates)
        ax.scatter(candidate_x, candidate_y, s=50, c='gray',
                   marker='o', alpha=0.5, label='候选点')

        # 绘制部署的中继
        if deployed_relays:
            relay_x, relay_y = zip(*deployed_relays)
            ax.scatter(relay_x, relay_y, s=150, c='green',
                       marker='D', label='部署中继', zorder=5)

            # 标注中继编号
            for i, (x, y) in enumerate(deployed_relays):
                ax.text(x, y + 1.5, f'R{i + 1}', ha='center', fontweight='bold')

        # 绘制通信范围
        for relay in deployed_relays:
            circle = plt.Circle(relay, 20, color='green', alpha=0.1)
            ax.add_patch(circle)

        ax.set_xlabel('X坐标 (km)')
        ax.set_ylabel('Y坐标 (km)')
        ax.set_title('中继无人机部署方案')
        ax.legend()
        ax.grid(True, alpha=0.3)
        ax.set_aspect('equal')
        plt.tight_layout()
        plt.show()


def main():
    """主执行程序"""
    print("=" * 50)
    print("       无人机森林消防设备配置优化系统")
    print("=" * 50)

    # 设置随机种子确保结果可重现
    np.random.seed(42)

    try:
        # 初始化模型
        model = WildfireEquipmentModel()

        # 1. 基础年分析
        print("\n1. 基础年分析 (第0年)")
        events = model.generate_fire_events(year=0)
        events_with_demand = model.calculate_equipment_demand(events)

        # 显示前5次火灾信息
        print("\n前5次火灾事件详情:")
        for i, event in enumerate(events_with_demand[:5]):
            print(f"  火灾{i + 1}: 区域{event['region'] + 1}, 面积{event['area']}km², "
                  f"小队{event['squads']}个, 需要SSA{event['n_SSA']}架, "
                  f"中继{event['n_R']}架")

        # 时间序列分析
        D_SSA, D_R = model.create_demand_timeseries(events_with_demand)

        # 设备配置优化
        config = model.optimize_equipment_config(np.max(D_SSA), np.max(D_R))

        print(f"\n基础年优化配置:")
        print(f"  SSA无人机: {config['x_SSA_opt']} 架")
        print(f"  中继无人机: {config['x_R_opt']} 架")
        print(f"  总投资: ${config['total_cost']:,.0f}")
        print(f"  峰值需求 - SSA: {config['peak_demand_SSA']:.1f}, "
              f"中继: {config['peak_demand_R']:.1f}")

        # 2. 多年扩容分析
        print("\n2. 多年扩容分析 (10年规划)")
        multi_year_df = model.multi_year_expansion(config, years=10)

        # 显示多年分析摘要
        print("\n多年扩容分析摘要:")
        print(multi_year_df[['year', 'inventory_SSA', 'inventory_R', 'annual_cost']].round(0))

        # 3. 中继部署优化
        print("\n3. 中继部署优化")
        relay_optimizer = RelayDeploymentOptimizer()
        eoc, frontlines, candidates = relay_optimizer.generate_scenario()
        relay_result = relay_optimizer.optimize_relay_deployment(eoc, frontlines, candidates)

        print(f"中继部署结果:")
        print(f"  需要部署: {relay_result['num_relays']} 架中继无人机")
        print(f"  部署位置: {relay_result['deployed_positions']}")

        # 4. 预算报告
        print("\n4. 预算报告")
        total_10yr_cost = multi_year_df['annual_cost'].sum()
        avg_annual_cost = multi_year_df['annual_cost'][1:].mean()

        print(f"  首年投资: ${config['total_cost']:,.0f}")
        print(f"  10年总投资: ${total_10yr_cost:,.0f}")
        print(f"  后续年度平均投资: ${avg_annual_cost:,.0f}")

        # 5. 可视化
        print("\n5. 生成可视化图表...")
        Visualization.plot_demand_timeseries(D_SSA, D_R)
        Visualization.plot_multi_year_analysis(multi_year_df)
        Visualization.plot_relay_deployment(eoc, frontlines, candidates,
                                            relay_result['deployed_positions'])

        print("\n" + "=" * 50)
        print("程序执行完成！")
        print("=" * 50)

    except Exception as e:
        print(f"程序执行出错: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()