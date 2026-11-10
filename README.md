# EnglishLearner 教学资料库

英语课程大纲、教案、课件、课堂游戏、复习讲义与招生海报的归档目录。

## 课程大纲与教案

- [`lesson-plans/our-world-starter-phonics-24-lessons.docx`](lesson-plans/our-world-starter-phonics-24-lessons.docx)
  Our World Starter × Phonics 1 故事化启蒙课程（24 课修订版）：L1 直接进入 Unit 1，短元音前置，L2 开始 CVC，并采用有情节的绘本。
- [`lesson-plans/our-world-starter-lesson1-my-school.pptx`](lesson-plans/our-world-starter-lesson1-my-school.pptx)
  Starter 第 1 课 `My School!`（53 页多感官照片版）：真实感照片词卡（AI 生成原创素材）替代矢量图；每个单词一个专属感官入口 + 动作 chant（翻书声、慢速揭示、微距猜物、空中画彩虹、纸飞机、白板魔术、敲击乐、音乐椅子）；每 2 词插入复习游戏（回声山、喂怪兽、心灵感应、唇语）+ 声音袋 + 闪卡 + 递进式 What's Missing；句型与 Phonics 游戏含情绪骰子、Peek-a-boo、猜卡、Pass the Bag、Freeze Dance、Pop or Bounce、Swat It、Robot Talk、触卡接力；每个词 8-10 次跨感官接触 + Day1/Day3/Day7 回家温习计划；内嵌 4 段 Kevin MacLeod 配乐。
- [`lesson-plans/our-world-1-phonics-24-lessons.docx`](lesson-plans/our-world-1-phonics-24-lessons.docx)
  Our World 1 × Phonics 课程大纲。
- [`lesson-plans/our-world-starter-phonics-syllabus.docx`](lesson-plans/our-world-starter-phonics-syllabus.docx)
  Starter 早期版本，保留作历史参考。
- [`lesson-plans/phonics-teacher-handbook.docx`](lesson-plans/phonics-teacher-handbook.docx)
  Phonics 教师手册。
- [`lesson-plans/fricatives-teacher-notes.docx`](lesson-plans/fricatives-teacher-notes.docx)
  摩擦音教师笔记。
- `lesson-plans/demo-lesson-*.docx` / `lesson-plans/demo-slides-*.pptx`
  Starter 与 Level 1 玩具主题试听课教案和 PPT。

## 阅读与语法课件

`decks/` 中保存可直接在浏览器打开的 HTML 课件：

- `unit3-the-challenge-of-alzheimers.html`：原版教材阅读课。
- `reading-unit2-identifying-trends.html`：趋势识别阅读课。
- `unit2-review-tenses.html`：时态复习。
- `lesson-20`、`lesson-21`、`lesson-24`、`lesson-48`–`lesson-52`：分课阅读材料。
- `phonics-fricatives-lab.html` / `fricatives-teacher-slides.html`：摩擦音互动课件。

## 课堂游戏

`games/` 包含浏览器互动游戏：

- `vowel-phonics-game.html`
- `plosives-game.html`
- `fricatives-game.html`
- `affricates-game.html`
- `decoder-game.html`

## 讲义与测评

`worksheets/` 包含：

- 六年级 Unit 1–3 期中复习讲义与教师答案。
- 时态、从句复习讲义与答案。
- KET–PET 分班测评。
- Sentence Builder 家庭作业。

## 海报

- [`kids-course-poster.png`](kids-course-poster.png)：可直接发布的 PNG。
- [`kids-course-poster.svg`](kids-course-poster.svg)：儿童课程海报 SVG 源文件。
- [`course-poster.svg`](course-poster.svg)：课程海报 SVG 源文件。

## 可复现生成

`scripts/build_starter_phonics_syllabus.py` 用于生成最新的 24 课 Starter 课程大纲。运行：

```bash
python3 scripts/build_starter_phonics_syllabus.py
```

`scripts/build_starter_lesson1_deck.py` 用于生成 Starter 第 1 课 PPTX。