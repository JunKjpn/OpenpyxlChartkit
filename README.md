# OpenpyxlChartKit

![Two-Dimensional Grouping](docs/images/two_dimensional_grouping.png)

**OpenpyxlChartKit** は、`openpyxl` を利用した Excel
グラフ作成を、宣言的（Declarative）に記述できる Python ライブラリです。

通常の `openpyxl` では `Chart`・`Series`・`Reference`
を個別に生成する必要がありますが、本ライブラリでは **ChartSpec**
を定義することで、グラフの種類や系列、グルーピング、レイアウトをまとめて指定できます。

------------------------------------------------------------------------

# 特徴

-   宣言的な API によるグラフ作成
-   Scatter Chart / Bar Chart に対応
-   縦持ち・横持ちデータに対応した Series Grouping
-   Horizontal Chart Grouping
-   Vertical Chart Grouping
-   Two-Dimensional Chart Grouping
-   グラフサイズに応じた自動レイアウト
-   openpyxl ベースの Excel 出力

------------------------------------------------------------------------

# 従来の openpyxl との違い

通常の `openpyxl` では、Chart・Reference・Series
の生成やレイアウト計算、グラフ配置などを個別に実装する必要があります。

OpenpyxlChartKit では、`ChartSpec`
を定義してレンダリングするだけでグラフを生成できます。

``` python
spec = ChartSpec(...)

ChartRenderer().render(workbook, spec)
```

**「どのように作るか」ではなく、「どんなグラフを作りたいか」**
を記述できることを目的として設計しています。

------------------------------------------------------------------------

# インストール

``` bash
git clone https://github.com/JunKjpn/OpenpyxlChartKit.git
cd OpenpyxlChartKit
pip install -e .
```

------------------------------------------------------------------------

# Quick Start

次の例では、既存の Excel ファイルを読み込み、散布図を作成します。

``` python
from openpyxl import load_workbook

from openpyxl_chartkit import (
    AxisSpec,
    ChartRenderer,
    ChartSpec,
    LayoutSpec,
    SeriesSpec,
)

workbook = load_workbook("sample.xlsx")

spec = ChartSpec(
    source_sheet="Sheet1",
    output_sheet="Sheet1",
    chart_type="scatter",
    x=AxisSpec(
        column="Strain %",
        title="Strain (%)",
    ),
    series=(
        SeriesSpec(
            y=AxisSpec(
                column="Stress MPa",
                title="Stress (MPa)",
            ),
        ),
    ),
    layout=LayoutSpec(
        width=20,
        height=15,
    ),
    title_template="Stress-Strain Chart",
)

ChartRenderer().render(workbook, spec)
workbook.save("output.xlsx")
```

`source_sheet` はデータを読み込むシート、`output_sheet`
はグラフを配置するシートです。同じシート名を指定すれば、データシートにグラフを追加できます。

------------------------------------------------------------------------

# 機能紹介

## Series Grouping

`SeriesSpec` の `grouping` を指定すると、1つの `SeriesSpec`
から複数の系列を自動生成できます。

Series Grouping には、データの形式に応じて `source="row"` と
`source="column"` の2種類があります。

### Row-based Series Grouping（縦持ちデータ）

同じ Y 列を、指定した列の値ごとに分割して系列を作成します。

例えば、次のようなデータを扱う場合です。

    SampleNum   Strain   Stress
  ----------- -------- --------
            1      0.1       10
            1      0.2       20
            2      0.1       30
            2      0.2       40

`SampleNum` の値ごとに `Stress` の系列を作成します。

``` python
SeriesSpec(
    y=AxisSpec("Stress"),
    grouping=GroupingSpec(
        columns=("SampleNum",),
        source="row",
    ),
)
```

`source` を省略した場合は `"row"` として扱われます。

### Column-based Series Grouping（横持ちデータ）

指定した列を、それぞれ独立した系列として扱います。

例えば、次のようなデータを扱う場合です。

    Strain   Sample A   Sample B   Sample C
  -------- ---------- ---------- ----------
         1         10         12         15
         2         20         22         25
         3         30         32         35

`Sample A`、`Sample B`、`Sample C` の3列から、それぞれ系列を作成します。

``` python
SeriesSpec(
    y=AxisSpec("Sample A"),
    grouping=GroupingSpec(
        columns=("Sample A", "Sample B", "Sample C"),
        source="column",
    ),
)
```

列名は Excel
シートのヘッダーと一致させてください。現時点では1行のヘッダーを前提としています。

![Series Grouping](docs/images/series_grouping.png)

------------------------------------------------------------------------

## Horizontal Chart Grouping

指定した列の値ごとにチャートを分割し、横方向へ並べます。

``` python
ChartSpec(
    horizontal_grouping=GroupingSpec(
        columns=("SampleNum",),
    ),
)
```

![Horizontal Grouping](docs/images/horizontal_grouping.png)

------------------------------------------------------------------------

## Vertical Chart Grouping

指定した列の値ごとにチャートを分割し、縦方向へ並べます。

``` python
ChartSpec(
    vertical_grouping=GroupingSpec(
        columns=("Cycle",),
    ),
)
```

![Vertical Grouping](docs/images/vertical_grouping.png)

------------------------------------------------------------------------

## Two-Dimensional Chart Grouping

Horizontal Chart Grouping と Vertical Chart Grouping
を組み合わせることで、2次元レイアウトを自動生成できます。

``` python
ChartSpec(
    horizontal_grouping=GroupingSpec(
        columns=("SampleNum",),
    ),
    vertical_grouping=GroupingSpec(
        columns=("Cycle",),
    ),
)
```

![Two-Dimensional Grouping](docs/images/two_dimensional_grouping.png)

------------------------------------------------------------------------

# レイアウトエンジン

グラフサイズやグルーピングによって必要な配置領域を計算し、複数のチャートが重ならないように配置します。

グラフのサイズは `LayoutSpec` で指定できます。

``` python
LayoutSpec(
    width=20,
    height=12,
)
```

複数の `ChartSpec`
を同じシートに順番に配置する場合は、`ChartLayoutSession` を共有します。

``` python
session = ChartLayoutSession()

renderer.render(workbook, first_spec, session=session)
renderer.render(workbook, second_spec, session=session)
```

------------------------------------------------------------------------

# サンプル

`examples` フォルダには、ライブラリの利用例を用意しています。

  Example              Description
  -------------------- ------------------
  `strain_stress.py`   応力-ひずみ曲線
  `time_stress.py`     時系列応力データ

サンプルデータは `samples/` に含まれています。

------------------------------------------------------------------------

# 設計思想

OpenpyxlChartKit は、次の3点を重視しています。

### Declarative

実装ではなく、グラフの仕様を書く。

### Minimal

最小限の API でグラフを作成できる。

### Composable

Series・Grouping・Layout を組み合わせて利用できる。

------------------------------------------------------------------------

# 実装済み機能

-   ✅ Scatter Chart
-   ✅ Bar Chart
-   ✅ Row-based Series Grouping
-   ✅ Column-based Series Grouping
-   ✅ Horizontal Chart Grouping
-   ✅ Vertical Chart Grouping
-   ✅ Two-Dimensional Chart Grouping
-   ✅ Automatic Layout Engine

------------------------------------------------------------------------

# 今後追加を検討している機能

-   Pie Chart
-   Line Chart
-   Color Theme
-   Secondary Axis
-   Trendline
-   Legend Customization

------------------------------------------------------------------------

# ライセンス

MIT License
