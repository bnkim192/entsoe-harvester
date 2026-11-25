"""
폴란드 5년(2022-01-01 ~ 2025-12-31) 일별 가격·수요·풍력·태양광·기상 + EUA 경매가 한 번 받기
(2026-11-25 신규, 2022~2026 월별 가격 요인 분해용. 2026년은 pl_drivers_daily.json 에 이미 있음)

자료  fetch_pl_winter_hist.py 의 수집 함수를 그대로 쓰고 기간만 4년 전체로 바꾼다.
      ENTSO-E A44 가격 · A65 수요 · A75 전원별 발전 / Open-Meteo Archive 풍속·기온 / ECB EUR→PLN
      EUA = fetch_pl_market.get_eua() 를 시작일 2022-01-01 로 호출(EEX 연도별 경매 결과 파일)
산출  pl_hist_5y.json  {daily:[...], monthly:[...], eua:{날짜:값}}   ※ 공개 자료만, 회사 정보 없음
환경변수  ENTSOE_TOKEN (GitHub Secrets)
"""
import json
import fetch_pl_winter_hist as wh
import fetch_pl_market as fm

wh.PERIODS = [("2022-01-01", "2026-01-01")]
wh.OUT = "pl_hist_5y.json"


def main():
    wh.main()
    fm.START = "2022-01-01"
    try:
        eua = fm.get_eua()
    except Exception as e:
        wh.fd.log("   [EUA] 실패 %r" % e)
        eua = {}
    with open(wh.OUT, encoding="utf-8") as f:
        d = json.load(f)
    d["eua"] = {k: v for k, v in sorted(eua.items()) if k < "2026-01-01"}
    d["units"]["eua"] = "EUR/t (EEX 경매가, 같은 날 여러 경매는 평균)"
    d["collect_market"] = fm.LOG
    with open(wh.OUT, "w", encoding="utf-8") as f:
        json.dump(d, f, ensure_ascii=False, indent=1)
    wh.fd.log("[OK] %s · 일 %d · EUA %d일" % (wh.OUT, len(d.get("daily", [])), len(d["eua"])))


if __name__ == "__main__":
    main()