import folium
import json
import pandas as pd
from pathlib import Path


def create_japan_map_from_excel(excel_path, geojson_path, center=[36.2048, 138.2529], zoom_start=5):
    """
    ExcelファイルとGeoJSONファイルから日本地図を生成します。
    
    この関数は、Foliumを用いて日本地図を描画します。
    ・国土地理院のタイルを利用し、地図上に日本の領域と国外をマスクするレイヤーを追加します。
    ・Excelに記載された座標情報から、各地点にマーカー（必要に応じて円も）を表示します。
    
    パラメータ:
        excel_path (strまたはPath): マーカー情報が記載されたExcelファイルのパス。
            ※ 必須カラム：'lat'（緯度）、'lon'（経度）、'info'（ツールチップ用テキスト）
            ※ 任意カラム：'color'（マーカー色, デフォルト:'blue'）、'show_circle'（円表示フラグ）、'circle_color'（円の色, デフォルト:'blue'）
        geojson_path (strまたはPath): 日本領域のGeoJSONファイルのパス。
        center (list): 地図の中心座標（デフォルトは日本の中心付近）。
        zoom_start (int): 初期ズームレベル。
        
    戻り値:
        folium.Map: 作成したFoliumの地図オブジェクト。
    """
    # ベースとなるFoliumマップを作成（タイルは後から追加するためNone）
    fmap = folium.Map(
        location=center,
        zoom_start=zoom_start,
        tiles=None,             # デフォルトタイル無し
        control_scale=True,     # 地図にスケール表示を追加
        prefer_canvas=True,     # Canvasを優先して描画（パフォーマンス向上）
        zoom_animation=True,    # ズームアニメーションを有効化
        max_bounds=True,        # 地図の移動範囲を制限
        zoom_control=True       # ズームコントロールを有効化
    )

    # 地図の移動範囲（パン操作の制限）を設定（日本周辺に限定）
    fmap.options['maxBounds'] = [[20.0, 122.0], [46.0, 154.0]]
    fmap.options['minZoom'] = 5 # ズームアウトの最小値
    fmap.options['maxZoom'] = 12 # ズームインの最大値

    # 国土地理院提供のタイルをベースレイヤーとして追加
    folium.TileLayer(
        tiles="https://cyberjapandata.gsi.go.jp/xyz/std/{z}/{x}/{y}.png",
        attr="地図出典：国土地理院",
        name="国土地理院タイル",
        overlay=False,
        control=True
    ).add_to(fmap)

    # GeoJSONファイルを読み込み、日本の領域情報を取得
    with open(geojson_path, encoding="utf-8") as geo_file:
        japan_geo = json.load(geo_file)

    # 国外（日本以外の領域）をマスクするためのGeoJSONデータを作成
    # 外側のポリゴンは地球全体、内側の穴部分に日本の領域を指定
    mask_geojson = {
        "type": "Feature",
        "geometry": {
            "type": "Polygon",
            "coordinates": [
                [
                    [180, 90], [180, -90], [-180, -90], [-180, 90], [180, 90]
                ],
                # 日本領域の外枠（GeoJSONの最初のポリゴン座標を利用）
                japan_geo["features"][0]["geometry"]["coordinates"][0]
            ]
        }
    }
    folium.GeoJson(
        data=mask_geojson,
        name="マスク",
        style_function=lambda feature: {
            "fillColor": "gray",
            "color": "gray",
            "fillOpacity": 0.8,
            "weight": 0
        }
    ).add_to(fmap)

    # Excelファイルからマーカー情報を読み込み
    df = pd.read_excel(excel_path)

    # Excelの各行を処理し、マーカーおよび必要に応じて円を追加
    for index, row in df.iterrows():
        # 緯度、経度、表示するラベル（ツールチップ）を取得
        lat = row['lat']
        lon = row['lon']
        label = row['info']

        # マーカーの色を設定（Excelに指定がない場合はデフォルトの青）
        icon_color = row['color'] if ('color' in row and pd.notnull(row['color'])) else 'blue'

        # 円を描画するかどうかのフラグを取得（指定がなければFalse）
        show_circle = bool(row.get('show_circle', False))
        # 円の色を設定（指定がない場合はデフォルトの青）
        circle_color = row['circle_color'] if ('circle_color' in row and pd.notnull(row['circle_color'])) else 'blue'

        # 指定した座標にマーカーを追加（ツールチップにラベルを設定）
        folium.Marker(
            location=[lat, lon],
            tooltip=label,
            icon=folium.Icon(color=icon_color)
        ).add_to(fmap)

        # 円表示フラグがTrueの場合、指定位置に半径10kmの円を追加
        if show_circle:
            folium.Circle(
                location=[lat, lon],
                radius=10000,   # 半径（メートル単位）
                color=circle_color,
                fill=True,
                fill_opacity=0.1
            ).add_to(fmap)

    # レイヤーコントロールを追加し、ユーザが表示レイヤーの切り替えを可能にする
    # folium.LayerControl().add_to(fmap)

    return fmap


def main():
    """
    メイン関数
    
    ・スクリプトが存在するディレクトリ内にあるExcelファイルとGeoJSONファイルを読み込み、
      create_japan_map_from_excel関数により地図を生成します。
    ・生成された地図はHTMLファイルとして保存されます。
    """
    # 現在のスクリプトのディレクトリを取得
    base_dir = Path(__file__).parent
    # 入力ファイル（ExcelとGeoJSON）のパスを定義
    excel_path = base_dir / "points.xlsx"
    geojson_path = base_dir / "japan_boundary.geojson"
    # 出力ファイル（HTML）のパスを定義
    output_path = base_dir / "japan_map.html"

    # 日本地図を生成
    japan_map = create_japan_map_from_excel(excel_path, geojson_path)
    # 生成した地図をHTMLファイルとして保存
    japan_map.save(str(output_path))

    # 地図作成完了のメッセージを表示
    print(f"地図の作成が完了しました: {output_path}")


# スクリプトとして実行された場合にmain()を呼び出す
if __name__ == '__main__':
    main()
