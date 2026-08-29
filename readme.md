# PaperLex (PaperLexicon 论文词汇助手)

专为计算机与人工智能学术文献精读打造的极简桌面悬浮词典与术语沉淀工具。

**核心查词、词形还原、动词形态、同根派生拓展、LaTeX/表格渲染与 Excel 归档 100% 免费运行**，零配置开箱即用；DeepSeek API 仅作为深度学术消歧与算法机理推演的**可选扩展**。

---
## 为什么用PaperLex?
* 1.市面主流查词软件太笨重,多广告,而**PaperLex仅40MB,支持划词查询,直接查询,英汉互译等多种功能**
* 2.查词软件只能在软件上复习生词,甚至要求会员专享,而**PaperLex本地一键Excel导出,便于整理复习,不受任何软件限制**
* 3.常规软件窗口过大,不能适配论文阅读环境,而**PaperLex按F8一键呼出,窗口小巧,并有学术深度解析功能,完美适配论文阅读场景**

---

##  下载链接(仅40MB):
##  https://github.com/Sheldon-Tan/PaperLex/releases/download/v4.1/PaperLex.zip
#说明:
* 1.如果需要deepseek接入,请点击"设置",配置api-key(也可以直接在可执行程序同目录下生成的config.json配置)
* 2.本项目完全开源(代码如上),不含病毒,如果被windows defender阻止,保留并运行即可
---

##  核心免费功能（零配置 · 开箱即用）

无需注册任何账号或配置 API Key，启动即可完整使用以下所有功能：

### 1. 划词/输入查询,自动转换大小写,自动纠错
* **F8 全局快捷划词**：在双栏 PDF 阅读器、浏览器或编辑器中划选文本，按 `F8` 秒级呼出置顶悬浮窗；未划选时直接激活并聚焦输入框。
* **句首大写智能转换**：划选句首大写的 `Intermediate`、`Linear`、`Attention` 时，自动绕过词典的孤立人名干扰，优先展示标准学术释义与词性。
* **自造复合词拆解**：针对文献中带连字符的复合词（如 `cross-attention`、`speaker-aware`），自动拆分多词素并提供分词详解。
* **模糊拼写纠错**：选中文本残缺或输入错误时，自动匹配并推荐相近候选词卡片。

<table>
  <tr>
    <td align="center">
      <img src="https://github.com/user-attachments/assets/33e888c2-f501-4955-b010-65358534c447" width="300" /><br>
      图 1：模糊拼写纠错
    </td>
    <td align="center">
      <img src="https://github.com/user-attachments/assets/13b36d11-9842-4dde-8dd4-cb40b3d9a86a" width="300" /><br>
      图 2：正常查词页面
    </td>
  </tr>
</table>
<table>
  <tr>
    <td align="center">
      <img src="https://github.com/user-attachments/assets/f99bf051-9f5c-4016-9172-4250346ec899" width="300" /><br>
      图 3：自动大小写转换推荐
    </td>
    <td align="center">
      <img src="https://github.com/user-attachments/assets/ed617568-921a-46cf-9df9-df312836a197" width="300" /><br>
      图 4：自造复合词拆解
    </td>
  </tr>
</table>



### 2. 提供动词各种变形
* **学术不规则原词逆推**：自动识别分词（`-ing`）、过去式/分词（`-ed`）、复数（`-es`/`-s`）以及学术论文中高频出现的拉丁/希腊借词不规则变化（如 `criteria -> criterion`、`media -> medium`），并在顶部提供一键跳转原词。
* **动词全形态与英/美拼写标注**：完整展示动词过去式、过去分词、现在分词及第三人称单数，自动辨析标注双写 `l`、`-ise/-ize`、`-our/-or` 等英美拼写差异。

### 3. 结构化联想记忆
* **同源词与派生网络**：自动提取词根派生词，按名词（`[n.]`）、形容词（`[adj.]`）、副词（`[adv.]`）分类排版呈现。
* **近义词与写作替换**：聚合学术同义词组，辅助论文阅读理解与论文写作润色。

### 4. 高清 LaTeX 公式与 Markdown 原生排版
* **300 DPI 超采样数学公式**：内置 Matplotlib MathText 矢量光栅化引擎（无需安装庞大的 TeX 环境），将行内变量 `\(` 与独立公式 `\[` 渲染为透明背景的高清抗锯齿图像。
* **原生嵌入式 GUI 表格**：自动将 Markdown 表格解析并渲染为带表头高亮、斑马纹条纹的原生 GUI 网格组件。
* **算法代码块容器**：伪代码与算法流程自动置入 Consolas 等宽字体的代码高亮容器中。
* **层级富文本解析**：完整支持 1~4 级标题、加粗文本、引用块与分割线排版。

### 5. 本地归档与智能去重
* **Excel 归档**：按 `Enter` 键即时将词条、音标、词性、完整释义与记录时间追加至本地 `words.xlsx`。
* **智能覆盖去重**：默认开启自动去重，重复保存已有词条时，自动用最新记录覆盖旧行并更新时间戳，确保生词本始终精炼整洁。
<p align="center">
  <table>
    <tr>
      <td align="center">
        <img src="https://github.com/user-attachments/assets/e6314430-9232-444d-898d-269cedae7faf" height="500" /><br>
        <sub>图 5：点击保存,Excel归档</sub>
      </td>
    </tr>
  </table>
</p>


### 6.牛津词典在线翻译
* **权威查询英汉双解**:自由切换"网页查询"和"牛津查询",包含权威释义,英汉双解,地道例句,提升对论文表达的理解度。
<p align="center">
  <table>
    <tr>
      <td align="center">
        <img src="https://github.com/user-attachments/assets/c0bcea25-1a2b-4ccc-8146-147ad37b9c6c" width="300" /><br>
        <sub>图 6：牛津双解查词界面</sub>
      </td>
    </tr>
  </table>
</p>


##  可选进阶功能：DeepSeek 学术深度推演（按需开启）

在设置中填入 API Key 即可解锁大模型驱动的前沿学术解析能力：

* **两阶段学术消歧嗅探（极速省 Token）**：遇到 `ASD`、`NMS`、`VAD` 等缩写时，大模型先以 ~50 Token 快速检测是否存在多个学术分支，自动弹出交互卡片（如 `[1] Active Speaker Detection (音视频多模态)` 与 `[2] Autism Spectrum Disorder (医学AI)`），避免盲目生成导致 Token 浪费与领域跑偏。
* **定向顶会深度推演**：选定具体方向后，结合顶会论文语境生成技术输入输出模态、损失函数数学机理、经典基准数据集与 SOTA 基线对比表格。
* **汉译英学术表达**：输入中文术语，生成 5 组适用于顶会论文写作的地道学术表达。
* **熔断保护机制**：配备实时秒级计数器与 60 秒超时自动熔断保护。
<table>
  <tr>
    <td align="center">
      <img src="https://github.com/user-attachments/assets/9de4d76c-c7d3-4ca3-9675-c99b42176640" width="300" /><br>
      图 7：一词多义时主动提示,节省token
    </td>
    <td align="center">
      <img src="https://github.com/user-attachments/assets/ea90ed67-d5c6-47f2-b49d-b669d4ddf9bd" width="300" /><br>
      图 8：深入解析学术缩写
    </td>
  </tr>
</table>

---

##  快捷键与操作说明

| 按键 / 交互 | 作用域 | 功能说明 |
| :--- | :--- | :--- |
| **`F8`** | 全局系统 | 捕获选中文本并唤出悬浮窗；未划选时直接唤起并聚焦输入框 |
| **`Enter`** | 悬浮窗口 | 保存当前词条至 `words.xlsx`（开启去重时自动覆盖旧记录） |
| **`Esc`** | 悬浮窗口 | 隐藏悬浮窗口并释放焦点 |
| **`查词`** | 悬浮窗口 | 检索输入框中的自定义英文单词、复合短语或中文 |
| **`查原词`** | 悬浮窗口 | 命中词形变化或句首大写词时，一键切换查询原型词 |
| **`深度解析`** | 悬浮窗口 | 触发两阶段学术消歧与 DeepSeek 大模型深度推演 |
| **`设置`** | 悬浮窗口 | 打开图形设置面板，配置 API Key 与自动去重开关 |

---

##  安装与快速开始

### 1. 环境依赖

环境要求：Python 3.8+  
系统支持：Windows 10 / 11

```bash
pip install requests keyboard pyperclip openpyxl matplotlib pillow
```

### 2. 启动应用
```bash
python PaperLex4.1.py
```
### 3. 配置说明（可选）
所有基础查词、动词变形、同源词网络、LaTeX 渲染与 Excel 归档功能无需任何配置即可直接使用。

若需开启 DeepSeek 学术深度推演，点击界面右下角 「设置」 按钮填入 API Key，或在生成的 config.json 中配置：

```bash
JSON
{
    "enabled": true,
    "api_key": "sk-your-deepseek-key-here",
    "base_url": "https://api.deepseek.com",
    "model": "deepseek-v4-flash",
    "auto_dedup": true
}
```
