import asyncio
from playwright.async_api import async_playwright
import pandas as pd

async def scrape_race_info(race_id: str):
    url = f"https://race.netkeiba.com/race/shutuba.html?race_id={race_id}"

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        page = await browser.new_page()
        await page.goto(url, timeout=60000)
        await page.wait_for_selector("tr.HorseList", timeout=30000)

        rows = await page.query_selector_all("tr.HorseList")

        horses = []
        for i, row in enumerate(rows):
            try:
                # 馬番
                uma_td = await row.query_selector("td[class^='Umaban']")
                uma_num = (await uma_td.inner_text()).strip() if uma_td else ""

                # 馬名
                name_a = await row.query_selector("span.HorseName a")
                name = (await name_a.inner_text()).strip() if name_a else ""

                # 単勝オッズ（Popularクラスで取得）
                odds_span = await row.query_selector("td.Popular span")
                odds = (await odds_span.inner_text()).strip() if odds_span else ""

                horses.append({
                    "馬番": uma_num,
                    "馬名": name,
                    "単勝オッズ": odds,
                    "勝率": ""
                })
            except Exception:
                continue

        await browser.close()

        # DataFrame に変換
        df = pd.DataFrame(horses)
        df["勝率"] = [f"=1/{len(df)}*100" for i in range(len(df))]
        df["期待値"] = [f"=C{i+2}*D{i+2}*100" for i in range(len(df))]
        df["期待値順位"] = [f"=RANK(E{i+2},E$2:E${len(df)+1},0)" for i in range(len(df))]

        # num = len(df)-2
        # df["勝率"] = [f"=1/{num}*100" for i in range(num)]
        # df["期待値"] = [f"=C{i+2}*D{i+2}*100" for i in range(num)]
        # df["期待値順位"] = [f"=RANK(E{i+2},E$2:E${num+1},0)" for i in range(num)]

        output_file = f"期待値計算_{race_id}.xlsx"
        df.to_excel(output_file, index=False)
        print(f"✅ Excelファイル出力完了: {output_file}")

# 実行（例: race_id を指定）
race_id = "202510020211"  # ← 任意のレースIDに変更可能
asyncio.run(scrape_race_info(race_id))
