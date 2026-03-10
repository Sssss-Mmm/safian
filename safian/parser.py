import re

def extract_phone(text):
    phone_pattern = re.compile(r'01[016789][-.\s]?\d{3,4}[-.\s]?\d{4}')
    match = phone_pattern.search(text)
    if match:
        raw = match.group()
        clean = re.sub(r'[^0-9]', '', raw)
        if len(clean) == 11:
            return f"{clean[:3]}-{clean[3:7]}-{clean[7:]}", text.replace(raw, "")
        elif len(clean) == 10:
            return f"{clean[:3]}-{clean[3:6]}-{clean[6:]}", text.replace(raw, "")
        return raw, text.replace(raw, "")
    return "", text

def extract_qty(text):
    qty_pattern = re.compile(r'(\d+)\s*(개|박스|송이|세트|건|봉지|포|병|단|상자)')
    matches = list(qty_pattern.finditer(text))
    if matches:
        last_match = matches[-1]
        return last_match.group(1), text.replace(last_match.group(0), " ")
    return "1", text

def extract_address(text):
    # 한국 주소의 가장 큰 특징: 도/시 로 시작하여, 동/면/읍/로/길 로 끝나고 숫자가 나옴, 뒤에 상세주소(아파트명 등)
    # 넓은 매칭을 위해 상세주소 부분에 자주 쓰이는 단어 포함
    regions = ['서울', '부산', '대구', '인천', '광주', '대전', '울산', '세종', '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주']
    region_str = '|'.join(regions)
    
    # 1. 시/도 지역명으로 시작하는 전형적 주소
    pattern1 = re.compile(f'((?:{region_str}|[가-힣]{{2,4}}시|[가-힣]{{2,4}}도)\\s+[가-힣]+(?:시|군|구)?\\s+[가-힣\\d\\s\\-]+(?:동|면|읍|리|로|길)\\s*[\\d\\-~]*(?:\\s+[가-힣a-zA-Z\\d\\s]+(?:차|동|호|층|아파트|빌라|오피스텔|타운|맨션|빌딩|단지|센터|상가|푸르지오|자이|더샵|캐슬|아이파크|힐스테이트|어울림|리슈빌|센트럴|하이엔드|파크|스위첸|데시앙|베르디움|루원|시티)[가-힣A-Za-z\\d\\s]*)*(?:\\s*\\d+동\\s*)?(?:\\s*\\d+호\\s*)?)')
    match1 = pattern1.search(text)
    if match1:
        addr = match1.group(1).strip()
        addr = re.sub(r'\s+', ' ', addr)
        return addr, text.replace(match1.group(0), " ")
        
    # 2. 지역명 없이 바로 시/구/동으로 시작하는 주소 
    pattern2 = re.compile(r'([가-힣]+(?:시|군|구)\s+[가-힣\d\s\-]+(?:동|면|읍|리|로|길)\s*[\d\-~]*(?:\s+[가-힣a-zA-Z\d\s]+(?:차|동|호|층|아파트|빌라|오피스텔|푸르지오|자이|캐슬|아이파크|하이엔드)[가-힣A-Za-z\d\s]*)*(?:\s*\d+동\s*)?(?:\s*\d+호\s*)?)')
    match2 = pattern2.search(text)
    if match2:
        addr = match2.group(1).strip()
        addr = re.sub(r'\s+', ' ', addr)
        return addr, text.replace(match2.group(0), " ")

    return "", text

def extract_memo(text):
    sentences = re.split(r'[\.\n\t]+', text)
    memo_keywords = ['문앞', '경비실', '소화전', '배송전', '연락', '부재시', '파손', '조심히', '맡겨', '놓고', '택배함', '배송', '기사', '연락요망']
    
    memo = ""
    remaining_text = []
    
    for s in sentences:
        s = s.strip()
        if not s: continue
        is_memo = any(kw in s for kw in memo_keywords)
        # 길이가 15자 이하이면서 메모 키워드가 있거나, 보통 끝에 붙음
        if is_memo and not memo and len(s) < 20:
            memo = s
        else:
            remaining_text.append(s)
            
    return memo, " ".join(remaining_text)

def extract_orderer(text):
    # 명시적 라벨
    name_pattern = re.compile(r'(이름|주문자|수취인|성명|받는분|주문인)[:\s]*([가-힣]{2,4})')
    match = name_pattern.search(text)
    if match:
         return match.group(2), text.replace(match.group(0), "")
    
    words = text.split()
    for w in words:
        clean_word = re.sub(r'[^가-힣]', '', w)
        if 2 <= len(clean_word) <= 4 and clean_word not in ['주문', '부탁', '감사', '배송', '주세요', '입니다', '박스', '사과', '배즙']:
             return clean_word, text.replace(w, "", 1)
             
    return "", text

def parse_tab_separated_order(text):
    """
    엑셀 등 표(Tab으로 구분)에서 복사해온 데이터 전문 파싱 로직
    """
    result = {
        "partner": "",
        "orderer": "",
        "mid_recipient": "",
        "mobile": "",
        "address": "",
        "qty": "1",
        "memo": "",
        "product_hint": ""
    }
    
    # 빈칸 제외하고 리스트업
    parts = [p.strip() for p in text.split('\t') if p.strip()]
    
    # 1. 핸드폰 번호(mobile) 제일 먼저 찾기
    for p in list(parts):
        if re.search(r'01[016789][-.\s]?\d{3,4}[-.\s]?\d{4}', p):
            raw = re.search(r'01[016789][-.\s]?\d{3,4}[-.\s]?\d{4}', p).group()
            clean = re.sub(r'[^0-9]', '', raw)
            if len(clean) == 11:
                result["mobile"] = f"{clean[:3]}-{clean[3:7]}-{clean[7:]}"
            elif len(clean) == 10:
                result["mobile"] = f"{clean[:3]}-{clean[3:6]}-{clean[6:]}"
            else:
                result["mobile"] = raw
            parts.remove(p)
            break
            
    # 2. 주소(address) 찾기
    addr_keywords = ['서울', '경기', '인천', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주', '시 ', '도 ', '군 ', '구 ', '동 ', '읍 ', '면 ', '로 ', '길 ']
    for p in list(parts):
        # 주소는 길이가 길고, 숫자가 있으면서 시/도/군/동/로 등의 키워드를 포함
        if (any(kw in p for kw in addr_keywords) and len(p) > 8 and re.search(r'\d', p)) or \
           re.search(r'(동|면|읍|리|로|길)\s*[\d\-~]+', p) or \
           '아파트' in p or '푸르지오' in p or '자이' in p:
            result["address"] = p
            parts.remove(p)
            break
            
    # 3. 이름(orderer, mid_recipient) 찾기
    names = []
    for p in list(parts):
        clean_p = p.replace(" ", "")
        # 보통 이름은 2~4글자 한글
        if 2 <= len(clean_p) <= 4 and re.match(r'^[가-힣]+$', clean_p):
            names.append(p)
            parts.remove(p)
            
    if len(names) == 1:
        result["orderer"] = names[0]
        result["mid_recipient"] = names[0]
    elif len(names) >= 2:
        result["orderer"] = names[0]
        result["mid_recipient"] = names[1]
        if len(names) >= 3:
            result["partner"] = names[0] # 첫번째 이름이 상호(파트너)였을 가능성
            result["orderer"] = names[1]
            result["mid_recipient"] = names[2]
            
    # 4. 수량 찾기
    for p in list(parts):
        if re.match(r'^\d+$', p) or re.match(r'^\d+\s*(개|박스|송이|세트|건|봉지|포|병|단|상자)$', p):
            result["qty"] = re.sub(r'[^0-9]', '', p)
            parts.remove(p)
            break
            
    # 5. 남은 문자열 분배 (가장 긴 것이 상호 또는 상품명 힌트일 확률 높음)
    if parts:
        # 길이나 키워드로 대략 짐작
        for p in list(parts):
            if any(kw in p for kw in ['문앞', '경비실', '연락', '택배함', '부재시']):
                result["memo"] = p
                parts.remove(p)
                break
                
    if parts and not result["partner"]:
        result["partner"] = parts.pop(0)
        
    if parts:
        result["product_hint"] = parts.pop(0)
        
    if parts and not result["memo"]:
        result["memo"] = " ".join(parts)
        
    return result

def parse_order_text(raw_text):
    """
    모든 종류의 주문 텍스트(카톡, 엑셀표 복사 등)를 분석합니다.
    """
    text = raw_text.strip()
    
    # 탭 문자가 2개 이상 있으면 엑셀(스프레드시트)에서 복사해온 구조적인 데이터로 판단!
    if text.count('\t') >= 2:
        return parse_tab_separated_order(text)
        
    # 이하 카카오톡 자유텍스트용 파서 동작
    result = {
        "orderer": "",
        "mobile": "",
        "address": "",
        "qty": "1",
        "memo": "",
        "product_hint": ""
    }
    
    result["mobile"], text = extract_phone(text)
    result["address"], text = extract_address(text)
    result["qty"], text = extract_qty(text)
    result["memo"], text = extract_memo(text)
    result["orderer"], text = extract_orderer(text)
    
    # 남은 잡동사니 단어들을 상품명 힌트로 취합
    final_text = re.sub(r'[^\w\s가-힣a-zA-Z0-9]', ' ', text)
    final_text = re.sub(r'\s+', ' ', final_text).strip()
    final_text = re.sub(r'상품명|품명|주문상품', '', final_text).strip()
    
    result["product_hint"] = final_text

    return result
