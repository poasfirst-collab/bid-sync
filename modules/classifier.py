RULES = [
    (['ISMS-P'],                          'ISMS-P인증'),
    (['ISMS'],                            'ISMS인증'),
    (['모의해킹', '모의 해킹'],           '모의해킹'),
    (['취약점 점검', '취약점점검',
      '취약점 분석', '취약점분석',
      '기반시설', '기반 시설', 'CIIP'],   '기술점검'),
    (['개인정보'],                        '개인정보보호'),
    (['보안관제', '관제'],                '보안관제'),
    (['클라우드'],                        '클라우드보안'),
    (['AI 보안', 'AI보안'],               'AI보안'),
    (['공급망'],                          '공급망보안'),
    (['컨설팅', '정보보호'],              '보안컨설팅'),
]

def classify(text: str) -> str:
    if not text:
        return '기타'
    t = str(text).upper()
    for keywords, label in RULES:
        for kw in keywords:
            if kw.upper() in t:
                return label
    return '기타'
