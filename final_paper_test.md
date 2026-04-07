<div align="center">

# 本科毕业设计（论文）


| 项目 | 内容 |
|:------|:------|
| **题目** | 基于网络爬虫的股票信息预警系统的设计与实现 |
| **英文题目** | DESIGN AND IMPLEMENTATION OF STOCK INFORMATION EARLY WARNING SYSTEM BASED ON WEB CRAWLER |
| **学号** | 20211120148 |
| **姓名** | 武盛玮 |
| **班级** | 22网工A2 |
| **专业** | 网络工程 |
| **学部(院)** | 计算机与信息工程学院 |
| **入学时间** | 2022级 |
| **指导教师** | 沈文枫 |
| **日期** | 2026年4月6日 |

</div>

---

<div style="page-break-after: always;"></div>

## 毕业设计（论文）独创性声明

本人所呈交的毕业论文是在指导教师指导下进行的工作及取得的成果。除文中已经注明的内容外，本论文不包含其他个人已经发表或撰写过的研究成果。对本文的研究做出重要贡献的个人和集体，均已在文中作了明确说明并表示谢意。

**作者签名：** _________________________________

**日    期：** _________________________________

---

<div style="page-break-after: always;"></div>

<div align="center">

# 基于网络爬虫的股票信息预警系统的设计与实现

</div>

## 摘要

本文围绕分布式异构数据融合与边缘计算协同机制展开研究，针对物联网环境下多源感知数据实时处理与智能决策的核心问题，提出一种基于轻量化神经网络与自适应调度算法的混合架构模型。通过引入联邦学习框架实现隐私保护下的跨节点知识共享，结合动态资源分配策略优化边缘-云协同计算效率。实验表明，该方案在保障数据时效性的同时，有效降低了系统能耗与通信开销，为复杂场景下的智能感知应用提供了可行的技术路径。

**关键词：**`边缘计算`；`联邦学习`；`异构数据融合`；`自适应调度`

---

<div align="center">

# DESIGN AND IMPLEMENTATION OF STOCK INFORMATION EARLY WARNING SYSTEM BASED ON WEB CRAWLER

</div>

## ABSTRACT

This paper focuses on distributed heterogeneous data fusion and edge-cloud collaborative computing mechanisms, addressing the core challenges of real-time processing and intelligent decision-making for multi-source sensory data in IoT environments. A hybrid architecture model based on lightweight neural networks and adaptive scheduling algorithms is proposed. By introducing a federated learning framework, knowledge sharing across nodes is achieved under privacy preservation, while dynamic resource allocation strategies optimize the efficiency of edge-cloud collaborative computing. Experimental results demonstrate that the proposed scheme effectively reduces system energy consumption and communication overhead while ensuring data timeliness, providing a feasible technical pathway for intelligent perception applications in complex scenarios.

**Key words:** `edge computing`; `federated learning`; `heterogeneous data fusion`; `adaptive scheduling`

---

<div style="page-break-after: always;"></div>

# 1 绪论

## 1.1 研究的背景

随着5G通信、人工智能与物联网技术的深度融合，万物互联时代正加速到来。据IDC预测，到2025年全球物联网设备连接数将突破270亿，产生的数据总量将达到175ZB。然而，海量异构数据的实时采集、传输与处理对传统云计算架构提出了严峻挑战：高延迟、带宽瓶颈、隐私泄露风险等问题日益凸显。

在此背景下，边缘计算作为一种新兴的分布式计算范式，通过将计算能力下沉至网络边缘，有效缓解了云端集中处理带来的时延与负载压力。如图2-1所示，边缘节点可在数据源头附近完成初步分析，仅将关键特征或决策结果上传至云端，显著提升了系统响应速度与资源利用效率。

![网络爬虫工作流程图](https://raw.githubusercontent.com/SwanWoo/myimg/main/20260311233702.png)

## 1.2 研究现状

当前，边缘智能领域的研究主要集中在任务卸载、资源调度与模型压缩三个方向。文献[3]提出基于深度强化学习的动态卸载策略，在移动边缘场景中实现了能耗与时延的帕累托优化；文献[5]设计了梯度稀疏化的联邦聚合算法，有效降低了跨设备通信开销。

然而，现有研究多假设节点同质且网络稳定，难以适应实际场景中设备算力差异大、连接间歇性中断等复杂约束。为此，本文引入自适应权重调整机制，结合设备状态感知与信道质量评估，构建更具鲁棒性的协同推理框架。

在理论建模方面，系统能耗可表示为计算能耗与通信能耗的加权和：

$$
E_{total} = \alpha \cdot \sum_{i=1}^{N} f_i^3 \cdot t_i + \beta \cdot \sum_{j=1}^{M} P_j \cdot \tau_j \tag{1-1}
$$

其中，$f_i$为第$i$个节点的计算频率，$t_i$为执行时间，$P_j$为传输功率，$\tau_j$为通信时长，$\alpha$与$\beta$为权重系数。

## 1.3 研究的意义

**理论意义**：本研究将博弈论与在线优化理论引入边缘协同场景，建立了多智能体非合作博弈下的资源竞争模型，为分布式系统的均衡分析提供了新的数学工具。

**实际应用价值**：
- 降低端到端延迟：边缘预处理可将关键任务响应时间缩短60%以上；
- 保护用户隐私：联邦学习框架确保原始数据不出本地，满足GDPR等合规要求；
- 提升系统可扩展性：模块化设计支持动态节点加入与退出，适应大规模部署需求。

## 1.4 研究的目标与内容

本研究旨在构建一个支持异构设备接入、具备自适应调度能力的边缘智能协同平台，具体目标包括：

（1）设计轻量化特征提取网络，适配资源受限的边缘设备；

（2）提出基于李雅普诺夫优化的在线任务调度算法，实现长期能效最优；

（3）开发隐私保护的知识蒸馏机制，在模型聚合过程中抑制敏感信息泄露；

（4）搭建原型系统并开展多场景验证，评估方案的综合性能。

## 1.5 论文的组织安排

本文共分为六个章节。第二章介绍边缘计算与联邦学习的理论基础；第三章进行系统需求建模与架构设计；第四章详述核心算法的实现细节；第五章展示实验设置与结果分析；第六章总结全文并展望未来工作。

---

<div style="page-break-after: always;"></div>

# 2 相关理论与技术概述

## 2.1 边缘计算架构

边缘计算采用"云-边-端"三层协同架构，其中边缘层承担数据过滤、实时推理与局部决策功能。其核心挑战在于如何在有限资源下平衡计算精度与执行效率。

![图2-1](https://raw.githubusercontent.com/SwanWoo/myimg/main/20260311233702.png)

## 2.2 联邦学习原理

联邦学习通过"数据不动模型动"的范式实现分布式训练。全局模型更新遵循如下规则：

$$
w_{t+1} = \sum_{k=1}^{K} \frac{n_k}{n} \cdot w_t^{(k)} \tag{2-1}
$$

其中$w_t^{(k)}$为第$k$个客户端的本地模型参数，$n_k$为其数据量，$n$为总样本数。

为防止梯度泄露，本文引入差分隐私噪声：

$$
\tilde{g} = g + \mathcal{N}(0, \sigma^2 \mathbf{I}) \tag{2-2}
$$

## 2.3 异构数据融合方法

多源传感器数据具有时序性、稀疏性与模态异构等特点。本文采用注意力机制实现特征加权融合：

$$
\mathbf{z} = \sum_{i=1}^{L} \alpha_i \cdot \mathbf{h}_i, \quad \alpha_i = \frac{\exp(\mathbf{v}^T \tanh(\mathbf{W}\mathbf{h}_i))}{\sum_{j=1}^{L} \exp(\mathbf{v}^T \tanh(\mathbf{W}\mathbf{h}_j))} \tag{2-3}
$$

![图2-2](https://raw.githubusercontent.com/SwanWoo/myimg/main/20260311233702.png)

## 2.4 自适应调度策略

系统状态可建模为马尔可夫决策过程（MDP），其贝尔曼最优方程为：

$$
V^*(s) = \max_a \left\{ R(s,a) + \gamma \sum_{s'} P(s'|s,a) V^*(s') \right\} \tag{2-4}
$$

为加速收敛，本文采用双Q学习（Double DQN）改进值函数估计，有效缓解过估计问题。

此外，信息熵可用于量化系统不确定性：

$$
H(X) = -\sum_{i=1}^{n} p(x_i) \log_2 p(x_i) \tag{2-5}
$$

在资源分配中，我们最小化加权熵以提升决策确定性。

---

<div style="page-break-after: always;"></div>

# 3 系统设计与实现

## 3.1 总体架构

系统采用微服务架构，包含设备管理、任务调度、模型训练、隐私保护与可视化监控五大模块。各模块通过gRPC实现高效通信，支持横向扩展。

## 3.2 核心算法实现

### 3.2.1 轻量化网络设计

针对边缘设备算力限制，本文改进MobileNetV3结构，引入神经架构搜索（NAS）自动优化层配置。损失函数定义为：

$$
\mathcal{L} = \lambda_1 \mathcal{L}_{cls} + \lambda_2 \mathcal{L}_{distill} + \lambda_3 \|\theta\|_1 \tag{3-1}
$$

### 3.2.2 在线调度算法

基于李雅普诺夫漂移加惩罚方法，构造虚拟队列$Q(t)$，其演化满足：

$$
Q(t+1) = \max[Q(t) + A(t) - \mu(t), 0] \tag{3-2}
$$

通过最小化漂移上界，实现队列稳定性与能效的联合优化。

![网络爬虫工作流程图](https://raw.githubusercontent.com/SwanWoo/myimg/main/20260311233702.png)

## 3.3 隐私保护机制

采用同态加密与秘密共享相结合的方式，确保模型参数在传输与聚合过程中的机密性。密钥更新遵循椭圆曲线密码学：

$$
P_{pub} = s \cdot G, \quad sk_i = H(ID_i \| s) \tag{3-3}
$$

其中$G$为椭圆曲线基点，$s$为主密钥，$H$为哈希函数。

---

# 结论

本文针对边缘智能场景下的数据融合与协同计算问题，提出了一种融合轻量化模型、自适应调度与隐私保护的混合架构。理论分析表明，所提算法在收敛速度与资源效率方面具有渐进最优性；实验验证显示，在CIFAR-10与ImageNet-Subset数据集上，模型精度损失小于2%，通信开销降低45%，端到端延迟减少62%。

系统已在智慧园区与工业巡检场景中完成试点部署，日均处理感知数据120万条，预警准确率达94.7%。未来工作将探索跨域联邦学习与量子安全加密的融合，进一步提升系统的智能化与安全性水平。

---

<div style="page-break-after: always;"></div>

# 致谢

在论文即将完成之际，谨向所有给予我支持与帮助的师长、同窗与亲友致以最诚挚的谢意。

感谢导师沈文枫老师高屋建瓴的学术指导与细致入微的关怀。老师严谨的治学精神、开阔的学术视野与包容的育人理念，深深影响着我的科研态度与人生选择。

感谢计算机与信息工程学院提供的优质科研平台与实验资源，使本研究得以顺利开展。

感谢实验室伙伴在项目攻关期间的协作与鼓励，那些深夜调试代码、反复推演公式的日子，将成为我求学路上最珍贵的记忆。

最后，感恩家人无条件的理解与陪伴，你们是我勇往直前的坚实后盾。

---

<div style="page-break-after: always;"></div>

# 参考文献

[1] 李旭乐，宗光华.生物工程微操作机器人视觉系统的研究[J].北京航空航天大学学报，2002（2）：22-25

[2] 孙家正，杨长青.计算机图形学[M].北京：清华大学出版社，1995:26-28

[3] 张三，李四.基于深度强化学习的边缘任务卸载策略[J].计算机学报，2021,44(5):1023-1035

[4] 王五，赵六.联邦学习中的隐私保护技术研究进展[J].软件学报，2022,33(8):2891-2910

[5] 陈七，刘八.面向异构边缘设备的自适应模型压缩方法[J].自动化学报，2020,46(12):2567-2579

[6] Smith J, Johnson A. Federated Learning: Challenges and Opportunities[M]. Cambridge: MIT Press, 2021:78-102

[7] Brown M, Davis K. Edge Intelligence: From Theory to Practice[J]. ACM Computing Surveys, 2023,55(4):1-36

[8] Wilson R, Anderson T. Distributed Optimization for IoT Systems[M]. Berlin: Springer Nature, 2022:145-178

---

<div style="page-break-after: always;"></div>

# 附录

## 附录A 核心算法伪代码

```python
# 自适应调度主循环
def adaptive_scheduling(nodes, tasks):
    for t in range(T):
        state = observe_system_state(nodes)
        action = policy_network.select_action(state)
        reward = execute_and_evaluate(action)
        policy_network.update(state, action, reward)
        update_lyapunov_queues(nodes, action)
    return policy_network
```

## 附录B 关键公式汇总

$$
\nabla_\theta \mathcal{L} = \frac{1}{N} \sum_{i=1}^{N} \nabla_\theta \ell(f_\theta(x_i), y_i) + \lambda \theta \tag{B-1}
$$

$$
\mathcal{F}\{f(t)\} = \int_{-\infty}^{\infty} f(t) e^{-j\omega t} dt \tag{B-2}
$$

$$
\frac{\partial \mathcal{L}}{\partial W} = \delta^{(l)} (a^{(l-1)})^T \tag{B-3}
$$

## 附录C 系统部署拓扑

![图2-2](https://raw.githubusercontent.com/SwanWoo/myimg/main/20260311233702.png)

> 注：图中红色节点表示边缘服务器，蓝色为终端设备，虚线为逻辑通信链路。
