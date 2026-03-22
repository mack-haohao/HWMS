"""
region_data.py — 省市区数据
直接读取 cpca 库内置的 adcodes.csv，无需联网
adcode 规则：前2位=省，前4位=市，前6位=区县
"""

from pathlib import Path

_provinces: list[str] = []
_cities: dict[str, list[str]] = {}
_districts: dict[str, list[str]] = {}


def _load() -> bool:
    global _provinces, _cities, _districts
    try:
        import cpca
        import pandas as pd
        csv_path = Path(cpca.__file__).parent / "resources" / "adcodes.csv"
        df = pd.read_csv(str(csv_path), dtype={"adcode": str})
        df["adcode"] = df["adcode"].str.zfill(12)
        df["pcode"]  = df["adcode"].str[:2]
        df["ccode"]  = df["adcode"].str[:4]

        prov_df = df[df["adcode"].str[2:]  == "0" * 10]
        city_df = df[(df["adcode"].str[4:] == "0" * 8) & (df["adcode"].str[2:4] != "00")]
        dist_df = df[(df["adcode"].str[6:] == "0" * 6) & (df["adcode"].str[4:6] != "00")]

        pcode2name = dict(zip(prov_df["pcode"], prov_df["name"]))
        ccode2name = dict(zip(city_df["ccode"], city_df["name"]))

        _provinces = list(prov_df["name"])
        _cities    = {p: [] for p in _provinces}
        _districts = {}

        for _, row in city_df.iterrows():
            pname = pcode2name.get(row["pcode"], "")
            if pname:
                _cities.setdefault(pname, []).append(row["name"])

        for _, row in dist_df.iterrows():
            cname = ccode2name.get(row["ccode"], "")
            if cname:
                _districts.setdefault(cname, []).append(row["name"])

        # 直辖市：把区直接挂在省下（跳过"市辖区"层）
        for dc in ["北京市", "上海市", "天津市", "重庆市"]:
            if dc in _cities:
                all_d = []
                for cname in _cities[dc]:
                    all_d.extend(_districts.get(cname, []))
                _cities[dc] = all_d

        return True
    except Exception as e:
        print(f"[region] 加载失败: {e}")
        return False


def _use_builtin():
    global _provinces, _cities, _districts
    _provinces = [
        "北京市","天津市","上海市","重庆市",
        "河北省","山西省","辽宁省","吉林省","黑龙江省",
        "江苏省","浙江省","安徽省","福建省","江西省","山东省",
        "河南省","湖北省","湖南省","广东省","海南省",
        "四川省","贵州省","云南省","陕西省","甘肃省","青海省",
        "内蒙古自治区","广西壮族自治区","西藏自治区",
        "宁夏回族自治区","新疆维吾尔自治区",
        "香港特别行政区","澳门特别行政区","台湾省",
    ]
    _cities = {}
    _districts = {}


def init_builtin():
    if not _load():
        _use_builtin()

def fetch_online() -> bool:
    return True

def init(online_callback=None):
    init_builtin()
    if online_callback:
        online_callback(True)

def get_provinces() -> list[str]:
    return _provinces

def get_cities(province: str) -> list[str]:
    return _cities.get(province, [])

def get_districts(city: str) -> list[str]:
    return _districts.get(city, [])
