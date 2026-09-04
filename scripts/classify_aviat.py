"""
Aviat 품목명/규격 텍스트에서 품목군 및 기술속성을 추출하는 규칙 기반 분류기.
- 문자열 유사도가 아닌 명시적 키워드/정규식 매칭만 사용한다 (임의 추정 금지 원칙).
- 이 모듈의 결과는 대부분 '추정' 수준이며, 원본에 명시된 표기만 확정으로 다룬다.
"""
import re

ITEM_GROUPS = [
    "Rack", "PDP/Power", "IDU/Chassis", "CPU/Control", "Fan", "Modem",
    "Interface Card", "ODU/RFU", "WTM", "WBX", "OBC2", "Duplexer",
    "Coupler/Hybrid", "Mount/Bracket", "Antenna", "Waveguide", "Cable",
    "Connector", "SFP/Optical Module", "Surge/Grounding", "License",
    "NMS/EMS", "PC/운영장비", "Tool", "기타",
]


def classify_item_group(name: str, spec: str) -> str:
    text = f"{name or ''} {spec or ''}"
    up = text.upper()

    def has(*keys):
        return any(k in up for k in keys)

    if has("_랙") or re.search(r"\bRACK\b", up):
        return "Rack"
    if has("전원분배반", "PDP") or re.search(r"\bPOWER SUPPLY\b", up) or has("전원 카드", "POE INJECTOR", "전원 및 데이터"):
        return "PDP/Power"
    if has("안테나") or "ANT-HP-DP" in up:
        return "Antenna"
    if re.search(r"\bWR\d+", up) or has("WAVEGUIDE", "W/G", "후렉시블", "엘립티컬", "ELLIPTICAL"):
        return "Waveguide"
    if has("LICENSE", "라이선스"):
        return "License"
    if re.search(r"컴퓨터|\bPC\(|NOTEBOOK|DESKTOP|노트북|CLIENT PC", up):
        return "PC/운영장비"
    if has("소프트웨어", "PROVISION PLUS", "NMS", " EMS ", "EMS 소프트웨어") or up.strip().endswith("EMS"):
        return "NMS/EMS"
    if re.search(r"\bFAN\b", up):
        return "Fan"
    if has("MOUNT", "BRACKET", "브라켓", "고정용", "HANGER", "마운팅"):
        return "Mount/Bracket"
    if has("SURGE", "낙뢰", "GROUND", "접지", "ARRESTOR"):
        return "Surge/Grounding"
    if has("DUPLEXER"):
        return "Duplexer"
    if has("COUPLER", "HYBRID", "커플러"):
        return "Coupler/Hybrid"
    if has("OBC2"):
        return "OBC2"
    if has("WBX"):
        return "WBX"
    if has("WTM"):
        return "WTM"
    if has("ODU", "OBU", "RFU", "TRANSCEIVER", "송수신부"):
        return "ODU/RFU"
    if has("SFP", "광모듈", "OPTIC GBIT MODULE", "OPTICAL MODULE"):
        return "SFP/Optical Module"
    if has("MODEM", "모뎀") or re.search(r"\bRAC\d*", up) or "RADIO ACCESS" in up:
        return "Modem"
    if has("MC-MV", "NCC", "NPC", "TERM-MV", "AUX-A", "제어 및", "관리 및 전원"):
        return "CPU/Control"
    if has("GBE", "STM1", "STM-1", "E1", "DAC", "MSE-A", "이더넷 카드", "인터페이스"):
        return "Interface Card"
    if has("CHASSIS", "IDU", "셀프", "SHELF"):
        return "IDU/Chassis"
    if has("CABLE", "케이블", "점퍼"):
        return "Cable"
    if has("CONNECTOR", "커넥터", "ADAPTER", "아답터", "콘넥터"):
        return "Connector"
    if has("TOOL", "공구"):
        return "Tool"
    return "기타"


def _freq_re(prefix: str) -> re.Pattern:
    """주파수 정규식 생성.
    좌우 경계에 '\\b' 대신 (?<![A-Za-z0-9])/(?![A-Za-z0-9])를 쓰는 이유:
    '\\b'는 '_'와 한글(밴드)을 모두 '단어문자'로 취급하므로 'HAX_U6GHz'의 언더스코어나
    'L6G밴드'의 'G'-'밴' 사이에서 경계를 인식하지 못해 주파수를 놓치거나 잘못 추출하는
    문제가 있었음(예: U6GHz -> 6GHz로 오추출). 영문자/숫자만 경계 판정에서 제외한다.
    """
    return re.compile(
        rf"(?<![A-Za-z0-9]){prefix}\s*G(?:Hz|hz|밴드)?(?![A-Za-z0-9])", re.IGNORECASE
    )


_FREQ_PATTERNS = [
    (_freq_re("L6"), "L6GHz"),
    (_freq_re("U6"), "U6GHz"),
    (_freq_re("111"), "__TYPO_111G__"),
    (_freq_re("11"), "11GHz"),
    (_freq_re("8"), "8GHz"),
    (_freq_re("6"), "6GHz"),
    (_freq_re("4"), "4GHz"),
]


def extract_freq(text: str):
    """주어진 텍스트에서 주파수 표기를 추출. (정규화값, 오타여부) 반환. 없으면 (None, False)."""
    if not text:
        return None, False
    for pat, label in _FREQ_PATTERNS:
        m = pat.search(text)
        if m:
            if label == "__TYPO_111G__":
                return "11GHz", True
            return label, False
    return None, False


def extract_band(text: str):
    if not text:
        return None
    m = re.search(r"\b(LOW|HIGH)\s*BAND\b", text, re.IGNORECASE)
    if m:
        return "Low Band" if m.group(1).upper() == "LOW" else "High Band"
    m = re.search(r"\b(LOW|HIGH)\b", text, re.IGNORECASE)
    if m:
        return "Low Band" if m.group(1).upper() == "LOW" else "High Band"
    return None


def extract_power_grade(text: str):
    if not text:
        return None
    m = re.search(r"ODU\s*\d+\s*(hp|sp|v2)\b", text)
    if m:
        tag = m.group(1)
        return {"hp": "High Power(hp)", "sp": "표준출력(sp)", "v2": "신형(v2)"}.get(tag)
    if re.search(r"\bHigh\s*Power\b", text, re.IGNORECASE):
        return "High Power"
    return None


def extract_config_notation(text: str):
    if not text:
        return None
    m = re.search(r"\b([2468])\s*\+\s*0\b", text)
    return f"{m.group(1)}+0" if m else None


def extract_channel_count(text: str):
    if not text:
        return None
    m = re.search(r"(\d+)\s*CH(?:A)?NNEL", text, re.IGNORECASE)
    return f"{m.group(1)}채널" if m else None


def extract_sd_flag(text: str):
    if not text:
        return False
    return bool(re.search(r"\bSD\b|SPACE\s*DIVERSITY|SD\s*OPTION|SD용|SD구성용", text, re.IGNORECASE))


def extract_ru(text: str):
    if not text:
        return None
    m = re.search(r"(\d)\s*RU\b", text)
    return f"{m.group(1)}RU" if m else None


def extract_ports(text: str):
    if not text:
        return None
    m = re.search(r"(\d+)\s*PORT", text, re.IGNORECASE)
    return f"{m.group(1)}Port" if m else None


def extract_interface_speed(text: str):
    if not text:
        return []
    tags = []
    up = text.upper()
    if re.search(r"\bE1\b|\d+E1\b", up):
        tags.append("E1")
    if "STM-1" in up or "STM1" in up or "OC3" in up:
        tags.append("STM-1")
    if re.search(r"\b10\s*GIGABIT|10GBE|XGBE\b", up):
        tags.append("10GbE")
    elif re.search(r"GIGABIT ETHERNET|\bGBE\b|1G OPTICAL|1G ELECTRICAL", up):
        tags.append("GbE")
    if "10/100/1000" in up:
        tags.append("10/100/1000BaseT")
    if "DS3" in up or "DS-3" in up or "E3" in up:
        tags.append("DS3/E3")
    return tags


def extract_cable_length(text: str):
    if not text:
        return None
    m = re.search(r"(\d+(?:\.\d+)?)\s*(mm|meter|m)\b", text, re.IGNORECASE)
    if not m:
        return None
    num, unit = m.group(1), m.group(2).lower()
    if unit == "mm":
        return f"{num}mm"
    return f"{num}m"


def extract_waveguide_spec(text: str):
    if not text:
        return None
    tokens = re.findall(r"\bWR\d+\b|\bUDR\d+\b|\bPDR\d+\b", text)
    return "/".join(dict.fromkeys(tokens)) if tokens else None


def extract_antenna_size(text: str):
    if not text:
        return None
    m = re.search(r"(\d+)\s*(FT|피트)\b", text, re.IGNORECASE)
    return f"{m.group(1)}ft" if m else None


def extract_polarization(text: str):
    if not text:
        return None
    if re.search(r"듀얼|이중극성|DUAL\s*POL", text, re.IGNORECASE):
        return "이중편파(Dual-Pol)"
    if re.search(r"SINGLE\s*POL|단일극성", text, re.IGNORECASE):
        return "단일편파(Single-Pol)"
    return None


_FAMILY_PATTERNS = [
    (re.compile(r"CTR\s*8312", re.IGNORECASE), "CTR8312"),
    (re.compile(r"CTR\s*8540", re.IGNORECASE), "CTR8540"),
    (re.compile(r"CTR\s*8740", re.IGNORECASE), "CTR8740"),
    (re.compile(r"WTM\s*4200", re.IGNORECASE), "WTM4200"),
    (re.compile(r"WTM\s*4500XT", re.IGNORECASE), "WTM4500XT"),
    (re.compile(r"WTM\s*4500", re.IGNORECASE), "WTM4500"),
    (re.compile(r"(?<![A-Za-z0-9])VR4(?![A-Za-z0-9])", re.IGNORECASE), "VR4"),
    (re.compile(r"(?<![A-Za-z0-9])VR10(?![A-Za-z0-9])", re.IGNORECASE), "VR10"),
    (re.compile(r"IDU[- ]?INUe", re.IGNORECASE), "INUe"),
    (re.compile(r"ODU\s*300", re.IGNORECASE), "ODU300"),
    (re.compile(r"ODU\s*600", re.IGNORECASE), "ODU600"),
]


def extract_family_key(text: str):
    """CTR8312/CTR8540처럼 주파수 태그가 없는 품목을 장비 계열(모델군)로 구분하기 위한 키.
    5단계(표준BoM 추정)에서 동일 이벤트에 서로 다른 계열의 무주파수 품목이 섞여 있을 때
    엉뚱하게 같은 구성으로 묶이지 않도록 하는 데 사용한다."""
    if not text:
        return None
    for pat, key in _FAMILY_PATTERNS:
        if pat.search(text):
            return key
    return None


def extract_all(name: str, spec: str) -> dict:
    combined = f"{name or ''} {spec or ''}"
    freq_name, typo_name = extract_freq(name)
    freq_spec, typo_spec = extract_freq(spec)

    issues = []
    if typo_name or typo_spec:
        issues.append("품명 또는 설명에 '111G' 형태 오타 의심(11GHz로 추정)")
    if freq_name and freq_spec and freq_name != freq_spec:
        issues.append(f"품명 주파수({freq_name})와 설명 주파수({freq_spec}) 불일치")

    freq_final = freq_name or freq_spec

    l6_u6 = None
    if freq_final == "L6GHz":
        l6_u6 = "L6GHz"
    elif freq_final == "U6GHz":
        l6_u6 = "U6GHz"

    return {
        "품목군": classify_item_group(name, spec),
        "장비계열": extract_family_key(combined) or "",
        "주파수_품명": freq_name or "",
        "주파수_설명": freq_spec or "",
        "주파수": freq_final or "",
        "L6_U6구분": l6_u6 or "",
        "LowHigh_Band": extract_band(combined) or "",
        "출력등급": extract_power_grade(combined) or "",
        "채널수": extract_channel_count(combined) or "",
        "구성표기": extract_config_notation(combined) or "",
        "SD_Diversity표기": "Y" if extract_sd_flag(combined) else "",
        "RU": extract_ru(combined) or "",
        "포트수": extract_ports(combined) or "",
        "인터페이스속도": "/".join(extract_interface_speed(combined)),
        "케이블길이": extract_cable_length(combined) or "",
        "Waveguide규격": extract_waveguide_spec(combined) or "",
        "안테나크기": extract_antenna_size(combined) or "",
        "편파": extract_polarization(combined) or "",
        "데이터품질이슈": "; ".join(issues),
    }
