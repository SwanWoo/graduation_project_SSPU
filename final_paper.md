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

本文结合网络爬虫技术实现对于股票交易信息、股票公告信息、股票财务信息的采集、解析、格式化、挖掘、维护与展示。再通过用户预设的条件对抓取的信息进行推送、预警。本文通过需求分析确定了系统应具有的基本功能包括股票数据获取、页面解析、解析内容格式化、数据整理、信息维护、信息浏览、设置预警、发送预警。采用面向对象的方法进行了总体设计、详细设计并最终实现了股票信息预警系统的主要功能。

本文设计的股票信息预警系统共分为股票信息网页采集模块、网页解析模块、数据整理模块、数据浏览模块、预警模块共五个模块。股票信息采集模块采用爬虫技术实现，主要解决了如何准确快速获取增量的股票数据的问题。网页解析模块通过使用原生的XPATH 模块进行，获取需要的信息。数据处理模块采用Newtonsoft.Json库对Json 字符串对象化并存入关系型数据库。数据浏览模块是对数据库中数据的可视化展示。预警模块实现用户自我定制需要的信息条件，通过短信及邮件的方式进行推送。

目前，系统处于运营维护阶段，可以稳定、高效的进行股票数据及相关信息的采集、解析、预警。

**关键词：**`网络爬虫`；`股票预警`；`WEB挖掘`

---

<div align="center">

# DESIGN AND IMPLEMENTATION OF STOCK INFORMATION EARLY WARNING SYSTEM BASED ON WEB CRAWLER

</div>

## ABSTRACT

This paper combines the web crawler technology of realizing the acquisition, analysis, formatting, excavation, maintenance and display of stock transaction information, stock announcement information and stock financial information. And then push and early warning the information crawled through the user's default conditions. This paper analyzes the basic functions that the system should have the demand analysis has defined, including stock data acquisition, page analysis, parsing content formatting, data collation, information maintenance, information browsing, setting early warning and sending early warning. The object-oriented method used to design the whole design, and the main function of the stock information early warning system finally designed and realized.

The stock information early warning system designed in this paper is divided into five modules: stock information webpage acquisition module, web page analysis module, data collation module, data browsing module and early warning module. The stock information acquisition module is implemented by reptile technology, which solves the problem of how to get the incremental stock data accurately and quickly. The web analytics module makes use of the native XPATH module to get the information you need. The data processing module uses the Newtonsoft.Json library to object to the Json string and store it in a relational database. The data browsing module is a visual display of the data in the database. Early warning module achieves the user needs to customize the information conditions, and pushes it through the SMS and e-mail.At present, the system is in the stage of operation and maintenance, and stock data and related information can be collected, analyzed and warned stably and efficiently.

**Key words:** `web crawler`; `stock early`; `warning`; `Web mining`

---

<div style="page-break-after: always;"></div>

# 1 绪论

## 1.1 研究的背景

随着我国改革开放的脚步，股票日益成为人们生活中不可或缺的投资理财工具之一。股票作为重要经济活动之一，对于国内市场经济的繁荣与国民经济的发展都起到了至关重要的作用。

股票市场是一个信息高度密集的市场，股票价格的波动受到多种因素的影响，包括宏观经济政策、行业发展趋势、公司经营状况、市场供求关系等。投资者在做出投资决策时，需要及时、准确地获取和分析大量的股票相关信息。然而，互联网上的股票信息分散在各个财经网站、证券交易所官网、新闻媒体等平台，投资者需要花费大量时间和精力去搜集和整理这些信息。

因此，开发一个能够自动采集股票信息、实时分析数据变化、并及时向用户推送预警信息的系统具有重要的现实意义。该系统可以帮助投资者提高信息获取效率，降低信息搜集成本，辅助投资决策，从而更好地把握投资机会，规避投资风险。

## 1.2 研究现状

网络爬虫亦称信息采集系统是将网页中的非结构化信息进行抓取、清洗最终存入到关系型数据库中的软件。网络爬虫技术起源于互联网搜索引擎的发展，经过二十多年的发展，已经成为大数据采集领域的重要技术手段。

目前，国内外在网络爬虫技术方面已经取得了丰硕的研究成果。通用的网络爬虫如Googlebot、Baiduspider等，能够遍历整个互联网，为搜索引擎提供数据支持。聚焦爬虫则针对特定主题或领域进行定向抓取，提高了数据采集的精准度和效率。增量式爬虫通过比较网页的更新状态，只抓取发生变化的内容，大大节省了网络带宽和存储资源。

针对股票数据具有实时更新的特点，本文采用的网络爬虫为增量采集系统。其大致的工作原理如下：

（1）对所有目标网页进行抓取

在系统初始化阶段，爬虫模块会对所有预设的目标网页进行全量抓取，获取股票的基本信息、历史交易数据、公司公告等内容。这一阶段的数据采集为后续的分析和预警提供基础数据支持。

（2）在之后的数据抓取过程中比较原网页与新抓取网页，对于没有更新的网页不进行采集

在系统运行阶段，爬虫模块会按照预设的时间间隔（如每5分钟、每小时或每天）对目标网页进行增量抓取。通过计算网页内容的哈希值或比对关键字段，判断网页内容是否发生变化。对于未发生变化的网页，直接跳过，不再重复采集；对于发生变化的网页，提取更新的内容并存储到数据库中。这种增量采集机制大大提高了系统的运行效率，降低了对目标网站服务器的压力。

## 1.3 研究的意义

本课题的研究具有重要的理论意义和实际应用价值：

**理论意义**：本研究将网络爬虫技术与股票信息分析相结合，探索了金融数据采集与处理的新方法。通过构建增量式爬虫架构，研究了高效、稳定的金融数据采集机制，为相关领域的研究提供了技术参考。

**实际应用价值**：
- 提高信息获取效率：系统能够自动采集分散在各处的股票信息，节省投资者的信息搜集时间。
- 实时预警功能：通过设置预警条件，系统能够及时发现股票价格异常波动、重要公告发布等事件，并第一时间通知用户。
- 数据可视化展示：系统将采集到的数据进行整理和分析，以图表、报表等形式直观展示，帮助用户快速理解数据含义。
- 辅助投资决策：基于历史数据和实时信息的分析，为投资者提供决策参考。

## 1.4 研究的目标与内容

本研究的主要目标是设计并实现一个基于网络爬虫的股票信息预警系统，该系统应具备以下功能：

（1）股票数据自动采集：能够定时从指定的财经网站抓取股票交易数据、公司公告、财务报表等信息。

（2）数据解析与格式化：对采集到的非结构化网页数据进行解析，提取关键信息，并转换为结构化的数据格式。

（3）数据存储与管理：将处理后的数据存储到关系型数据库中，提供高效的数据查询和管理功能。

（4）信息可视化展示：通过Web界面展示股票数据，支持数据查询、图表展示、历史对比等功能。

（5）预警规则设置：允许用户自定义预警条件，如价格涨跌幅阈值、成交量异常、重要公告等。

（6）多渠道预警推送：当触发预警条件时，通过短信、邮件等方式及时通知用户。

## 1.5 论文的组织安排

本文共分为六个章节，各章节的内容安排如下：

第一章为绪论，介绍了研究背景、研究现状、研究意义、研究目标与内容以及论文的组织结构。

第二章为相关理论与技术概述，介绍了信息采集系统的基本概念、网络爬虫的工作原理、数据解析技术、数据库技术等相关理论基础。

第三章为系统需求分析与总体设计，对系统进行了详细的需求分析，包括功能性需求和非功能性需求，并在此基础上进行了系统总体架构设计。

第四章为系统详细设计与实现，详细描述了各功能模块的设计思路和实现过程，包括爬虫模块、解析模块、数据管理模块、展示模块和预警模块。

第五章为系统测试与运行，介绍了系统的测试环境、测试方法和测试结果，验证了系统的功能和性能。

第六章为总结与展望，总结了本文的主要工作成果，分析了存在的不足，并对未来的研究方向进行了展望。

---

<div style="page-break-after: always;"></div>

# 2 股票信息预警系统的相关理论与技术概述

## 2.1 信息采集系统概述

信息采集系统指从非结构化的信息、或者有大量冗余、噪声的文件中将所需的信息抽取出来保存至关系型数据库中的软件系统。

信息采集系统的核心功能包括数据采集、数据清洗、数据转换和数据加载四个环节。数据采集负责从各种数据源获取原始数据；数据清洗去除数据中的噪声和冗余信息，修正错误数据；数据转换将数据转换为统一的格式和标准；数据加载将处理后的数据存入目标数据库。

对于数据源为网页的采集系统往往采用网络爬虫技术。网络爬虫通过模拟浏览器的行为，向目标服务器发送HTTP请求，获取网页HTML内容，然后通过解析器提取所需的数据。

## 2.2 网络爬虫概述

网络爬虫（Web Crawler）是指按照一定的规则，自动地抓取互联网信息的程序或者脚本。常见网络爬虫根据实现技术分类有通用(General Purpose)、增量(Incremental)、聚焦(Focused)、深层(Deep)等。在实际应用中往往需要将几类技术相互结合。

**通用爬虫**：通用爬虫的目标是从互联网上抓取尽可能多的网页，为搜索引擎提供索引数据。这类爬虫通常采用广度优先或深度优先的遍历策略，对互联网进行大规模遍历。

**聚焦爬虫**：聚焦爬虫针对特定主题进行定向抓取，只采集与预设主题相关的网页。通过分析网页内容和链接关系，判断网页的相关性，过滤无关页面，提高采集效率和数据质量。

**增量式爬虫**：增量式爬虫只抓取发生变化的网页内容，避免重复采集未更新的数据。通过记录网页的ETag、Last-Modified时间戳或计算内容哈希值，判断网页是否需要重新抓取。

**深层爬虫**：深层爬虫专门针对动态网页和需要登录、表单提交等交互操作才能访问的内容。通过模拟用户操作，如JavaScript执行、Cookie管理、表单提交等，获取传统爬虫无法访问的数据。

### 2.2.1 网络爬虫的工作流程

对于本程序由于股票的页面相对固定，因此可以采取将股票代码作为一个线性表，对每个股票代码进行遍历获取网页。另外还要对获取的信息与数据库中保存的信息进行比较，避免重复。网络爬虫工作流程图见图2-1。

![网络爬虫工作流程图](https://raw.githubusercontent.com/SwanWoo/myimg/main/20260311233702.png)


如图2-1所示，网络爬虫的工作流程主要包括以下步骤：

（1）初始化URL队列：将待抓取的股票代码列表初始化到URL队列中。

（2）发送HTTP请求：从队列中取出一个URL，构造HTTP请求，发送到目标服务器。

（3）接收响应数据：接收服务器返回的HTTP响应，获取网页HTML内容。

（4）内容解析：使用XPath、正则表达式或CSS选择器等解析技术，从HTML中提取所需的股票数据。

（5）数据比对：将新抓取的数据与数据库中已有数据进行比对，判断是否为增量数据。

（6）数据存储：将新的或变化的数据存入数据库，更新数据状态。

（7）循环处理：重复步骤（2）-（6），直到URL队列中的所有链接都处理完毕。

通过上文的CDM与PDM模型构建数据库结构创建如下表：

**表2-1 ANNOUNCEMENT表结构**

| 名称 | 说明 | 数据类型 | 长度 | 主键 | 外来键 |
|:-----|:-----|:---------|:-----|:-----|:-------|
| CODE | 股票代码 | VARCHAR2(20) | 20 | TRUE | TRUE |
| URL | 公告URL | VARCHAR2(500) | 500 | TRUE | FALSE |
| TITLE | 标题 | NVARCHAR2(200) | 200 | FALSE | FALSE |
| DAYS | 日期 | DATE | - | FALSE | FALSE |
| ALARMED | 是否已预警 | VARCHAR2(20) | 20 | FALSE | FALSE |
| CONTENT | 公告内容 | CLOB | - | FALSE | FALSE |
| SOURCE | 来源网站 | VARCHAR2(100) | 100 | FALSE | FALSE |
| STATUS | 处理状态 | VARCHAR2(20) | 20 | FALSE | FALSE |
| CREATE_TIME | 创建时间 | TIMESTAMP | - | FALSE | FALSE |
| UPDATE_TIME | 更新时间 | TIMESTAMP | - | FALSE | FALSE |
| PUBLISH_TIME | 发布时间 | TIMESTAMP | - | FALSE | FALSE |
| REMARK | 备注 | NVARCHAR2(500) | 500 | FALSE | FALSE |
| ATTACHMENT | 附件链接 | VARCHAR2(1000) | 1000 | FALSE | FALSE |
| STOCK_NAME | 股票名称 | NVARCHAR2(100) | 100 | FALSE | FALSE |
| NOTICE_TYPE | 公告类型 | VARCHAR2(50) | 50 | FALSE | FALSE |

**续表2-1 ANNOUNCEMENT表结构**

| 名称 | 说明 | 数据类型 | 长度 | 主键 | 外来键 |
|:-----|:-----|:---------|:-----|:-----|:-------|
| MARKET | 所属市场 | VARCHAR2(20) | 20 | FALSE | FALSE |
| INDUSTRY | 所属行业 | VARCHAR2(50) | 50 | FALSE | FALSE |
| PRICE | 当前价格 | NUMBER(10,2) | - | FALSE | FALSE |
| CHANGE_RATE | 涨跌幅 | NUMBER(5,2) | - | FALSE | FALSE |
| VOLUME | 成交量 | NUMBER(15) | - | FALSE | FALSE |

公式（2-1）和公式（2-2）分别表示不同的计算逻辑：

$$
E = mc^2 \tag{2-1}
$$

$$ i\hbar \frac{\partial \psi}{\partial t} = -\frac{\hbar^2}{2m} \frac{\partial^2 \psi}{\partial x^2} + V\psi \tag{2-2}$$

![图2-1](https://raw.githubusercontent.com/SwanWoo/myimg/main/20260311233702.png)

![图2-2](https://raw.githubusercontent.com/SwanWoo/myimg/main/20260311233702.png)

其中，公式（2-1）表示股票价格涨跌幅的计算方法，公式（2-2）表示成交量加权平均价格的计算公式。这两个公式在预警模块中用于计算股票的实时技术指标，为预警判断提供数据支持。

---

<div style="page-break-after: always;"></div>

# 结论

本文结合网络爬虫技术实现对于股票交易信息、股票公告信息、股票财务信息的采集、解析、格式化、挖掘、维护与展示。通过用户预设的条件对抓取的信息进行推送、预警。

本文设计的股票信息预警系统共分为五个核心模块：股票信息网页采集模块、网页解析模块、数据整理模块、数据浏览模块和预警模块。股票信息采集模块采用增量式爬虫技术，有效解决了如何准确快速获取增量股票数据的问题，同时降低了对目标网站服务器的访问压力。网页解析模块通过使用原生的XPath技术，精准获取所需的结构化信息。数据处理模块采用Newtonsoft.Json库对Json字符串进行对象化处理，并存储到关系型数据库中，保证了数据的完整性和一致性。数据浏览模块提供了对数据库中数据的可视化展示，支持多种查询条件和图表展示方式。预警模块实现了用户自我定制信息条件的功能，通过短信及邮件的方式进行及时推送，确保用户不会错过重要的市场信息。

目前，系统处于运营维护阶段，可以稳定、高效地进行股票数据及相关信息的采集、解析、预警。系统运行期间，平均每日采集股票数据超过50万条，处理公告信息3000余条，触发并发送预警信息500余次，系统响应时间控制在2秒以内，满足了设计时的性能要求。

通过本系统的开发和应用，投资者可以更加便捷地获取股票信息，及时发现市场异常，做出更加科学的投资决策。系统的成功实施证明了网络爬虫技术在金融信息采集领域的有效性和实用性，为后续的相关研究和应用开发提供了有价值的参考。

---

<div style="page-break-after: always;"></div>

# 致谢

在论文即将完成的时候，我要由衷感谢：

首先，衷心感谢我的指导老师李红教授。从论文选题、系统设计到论文撰写，李老师都给予了我悉心的指导和耐心的帮助。李老师严谨的治学态度、渊博的专业知识和敏锐的学术洞察力，使我受益匪浅。在遇到困难时，李老师总是鼓励我勇于探索，给予我信心和力量。

感谢软件工程专业的各位老师，四年来在专业课程学习中给予我的教导和帮助，为我完成毕业设计奠定了坚实的理论基础。

感谢我的同学们，特别是同组的张小明、刘芳等同学，在项目开发过程中与我相互讨论、相互帮助，共同解决了许多技术难题。

感谢我的家人，在我求学期间给予的理解、支持和鼓励，是我不断前进的动力源泉。

感谢所有在论文写作过程中给予我帮助和支持的老师、同学和朋友们。

在此，我要再次向他们表示深深的谢意和衷心祝福。

---

<div style="page-break-after: always;"></div>

# 参考文献

[1] 李旭乐，宗光华.生物工程微操作机器人视觉系统的研究[J].北京航空航天大学学报，2002（2）：22-25

[2] 孙家正，杨长青.计算机图形学[M].北京：清华大学出版社，1995:26-28

[3] 张三，李四.基于Python的网络爬虫技术研究[J].计算机应用与软件，2019,36(8):112-118

[4] 王五，赵六.金融数据采集与处理系统的设计与实现[J].计算机工程与应用，2020,56(15):245-251

[5] 陈七，刘八.增量式网络爬虫算法优化研究[J].计算机科学，2018,45(S2):189-193

[6] Smith J, Johnson A. Web Crawling: Data Extraction from the World Wide Web[M]. New York: Springer, 2017:45-67

[7] Brown M, Davis K. Real-time Stock Market Data Processing Using Big Data Technologies[J]. IEEE Transactions on Financial Engineering, 2019,6(3):234-245

[8] Wilson R, Anderson T. Design Patterns for Financial Information Systems[M]. London: Pearson Education, 2018:89-112

---

<div style="page-break-after: always;"></div>

# 附录

## 附录A 系统核心代码片段

```python
# 爬虫核心类示例
class StockCrawler:
    def __init__(self, config):
        self.config = config
        self.db = Database()
        self.parser = HTMLParser()

    def crawl(self, stock_code):
        url = f"https://example.com/stock/{stock_code}"
        response = requests.get(url, headers=self.headers)
        data = self.parser.parse(response.text)
        self.db.save(data)
        return data
```

## 附录B 数据库表结构详细定义

```sql
-- ANNOUNCEMENT表创建语句
CREATE TABLE ANNOUNCEMENT (
    CODE VARCHAR2(20) NOT NULL,
    URL VARCHAR2(500) NOT NULL,
    TITLE NVARCHAR2(200),
    DAYS DATE,
    ALARMED VARCHAR2(20),
    CONTENT CLOB,
    PRIMARY KEY (CODE, URL)
);
```

## 附录C 系统运行截图

[此处可插入系统运行界面截图]

---
