import os
import io
import re
import sys
import urllib.parse
import requests
import shutil
from PIL import Image

sys.stdout.reconfigure(encoding='utf-8')

btn_dir = os.path.join("images", "board_buttons")
os.makedirs(btn_dir, exist_ok=True)

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36'
}

def fetch_daum_sdb(name):
    url = f"https://search.daum.net/search?w=tot&q={urllib.parse.quote(name + ' 프로필')}"
    try:
        r = requests.get(url, headers=HEADERS, timeout=5)
        # SDB
        m = re.findall(r'fname=(https?%3A%2F%2F[^\s"\'&]*contentshub[^\s"\'&]*sdb[^\s"\'&]*)', r.text)
        if m: return urllib.parse.unquote(m[0])
        # Argon
        m = re.findall(r'fname=(https?%3A%2F%2Fsearch\d*\.kakaocdn\.net%2Fargon%2F[^&"\']+)', r.text)
        if m: return urllib.parse.unquote(m[0])
    except Exception as e:
        pass
    return None

def fetch_fandom(wiki, title):
    url = f"https://{wiki}.fandom.com/api.php?action=query&titles={title}&prop=pageimages&pithumbsize=600&format=json"
    try:
        r = requests.get(url, headers=HEADERS, timeout=5)
        pages = r.json().get('query', {}).get('pages', {})
        for pid, p in pages.items():
            if 'thumbnail' in p:
                return (p['thumbnail']['source'], f"https://{wiki}.fandom.com/")
    except Exception as e:
        pass
    return (None, None)

def fetch_wiki(title):
    try:
        r = requests.get(f'https://ko.wikipedia.org/w/api.php?action=query&titles={urllib.parse.quote(title)}&prop=pageimages&pithumbsize=600&format=json', headers=HEADERS, timeout=5)
        pages = r.json().get('query', {}).get('pages', {})
        for pid, p in pages.items():
            if 'thumbnail' in p:
                return p['thumbnail']['source']
    except:
        pass
    try:
        r = requests.get(f'https://en.wikipedia.org/w/api.php?action=query&titles={urllib.parse.quote(title)}&prop=pageimages&pithumbsize=600&format=json', headers=HEADERS, timeout=5)
        pages = r.json().get('query', {}).get('pages', {})
        for pid, p in pages.items():
            if 'thumbnail' in p:
                return p['thumbnail']['source']
    except:
        pass
    return None

def save_clean_img(src_url, dest_path, referer=None):
    try:
        h = {'User-Agent': HEADERS['User-Agent']}
        if referer: h['Referer'] = referer
        ir = requests.get(src_url, headers=h, timeout=6)
        if ir.status_code == 200:
            im = Image.open(io.BytesIO(ir.content))
            bg = Image.new('RGB', im.size, (255, 255, 255))
            if im.mode == 'RGBA':
                bg.paste(im, mask=im.split()[3])
            else:
                bg.paste(im.convert('RGB'))
            bg.thumbnail((600, 600))
            bg.save(dest_path, 'JPEG', quality=90)
            print(f"  Saved -> {dest_path} ({im.size})")
            return True
    except Exception as e:
        print(f"  Err saving {dest_path}: {e}")
    return False

# 1. anime_l1: 나루토 (기존 파일 확인)
p = os.path.join(btn_dir, "anime_l1.jpg")
if not os.path.exists(p) or os.path.getsize(p) < 5000:
    u, ref = fetch_fandom("naruto", "Naruto_Uzumaki")
    save_clean_img(u, p, ref)
print("anime_l1 (나루토): Ready")

# 2. anime_l2: 사스케 (유명 서브 주인공)
p = os.path.join(btn_dir, "anime_l2.jpg")
u, ref = fetch_fandom("naruto", "Sasuke_Uchiha")
if u: save_clean_img(u, p, ref)
print("anime_l2 (사스케): Ready")

# 3. anime_l3: 히소카 (헌터x헌터 고난도 매니아 캐릭터)
p = os.path.join(btn_dir, "anime_l3.jpg")
u, ref = fetch_fandom("hunterxhunter", "Hisoka_Morow")
if u: save_clean_img(u, p, ref)
print("anime_l3 (히소카): Ready")

# 4. actor_l1: 로버트 다우니 주니어
p = os.path.join(btn_dir, "actor_l1.jpg")
if not os.path.exists(p) or os.path.getsize(p) < 5000:
    u = fetch_wiki("Robert_Downey_Jr.")
    if u: save_clean_img(u, p)
print("actor_l1 (로다주): Ready")

# 5. actor_l2: 성동일 (명품 조연)
p = os.path.join(btn_dir, "actor_l2.jpg")
u = fetch_daum_sdb("성동일")
if u: save_clean_img(u, p)
print("actor_l2 (성동일): Ready")

# 6. actor_l3: 김의성 (개성파/고난도 씬스틸러)
p = os.path.join(btn_dir, "actor_l3.jpg")
u = fetch_daum_sdb("김의성")
if u: save_clean_img(u, p)
print("actor_l3 (김의성): Ready")

# 7. music_l1: 싸이 (초특급 국민가수)
p = os.path.join(btn_dir, "music_l1.jpg")
u = fetch_daum_sdb("싸이")
if u: save_clean_img(u, p)
print("music_l1 (싸이): Ready")

# 8. music_l2: 자이언티 (개성파 유명 보컬)
p = os.path.join(btn_dir, "music_l2.jpg")
u = fetch_daum_sdb("자이언티")
if u: save_clean_img(u, p)
print("music_l2 (자이언티): Ready")

# 9. music_l3: 양준일 (고난도/마니아)
p = os.path.join(btn_dir, "music_l3.jpg")
u = fetch_daum_sdb("양준일")
if u: save_clean_img(u, p)
print("music_l3 (양준일): Ready")

# 10. youtube_l1: 침착맨 (사용자 명시적 요청!)
p = os.path.join(btn_dir, "youtube_l1.jpg")
chim_src = os.path.join("images", "유튜버_1단계", "침착맨 (이말년).jpg")
if os.path.exists(chim_src):
    shutil.copyfile(chim_src, p)
    print("youtube_l1 (침착맨): Copied from 유튜버_1단계/침착맨 (이말년).jpg")
else:
    u = fetch_daum_sdb("침착맨")
    if u: save_clean_img(u, p)
    print("youtube_l1 (침착맨): Ready from Daum")

# 11. youtube_l2: 김풍 (침착맨 패밀리/유명 서브 크리에이터)
p = os.path.join(btn_dir, "youtube_l2.jpg")
u = fetch_daum_sdb("김풍")
if u: save_clean_img(u, p)
print("youtube_l2 (김풍): Ready")

# 12. youtube_l3: 충주맨 (김선태 주무관 / 공공 크리에이터 고난도)
p = os.path.join(btn_dir, "youtube_l3.jpg")
u = fetch_daum_sdb("김선태 충주맨") or fetch_daum_sdb("김선태")
if u: save_clean_img(u, p)
print("youtube_l3 (충주맨): Ready")

# 13. general_l1: 스티브 잡스 (초특급 글로벌 혁신가)
p = os.path.join(btn_dir, "general_l1.jpg")
u = fetch_wiki("Steve_Jobs")
if u: save_clean_img(u, p)
print("general_l1 (스티브 잡스): Ready")

# 14. general_l2: 박찬호 (코리안 특급 레전드)
p = os.path.join(btn_dir, "general_l2.jpg")
u = fetch_daum_sdb("박찬호")
if u: save_clean_img(u, p)
print("general_l2 (박찬호): Ready")

# 15. general_l3: 이창호 (바둑계의 살아있는 전설, 고난도)
p = os.path.join(btn_dir, "general_l3.jpg")
u = fetch_daum_sdb("이창호 기사") or fetch_daum_sdb("이창호")
if u: save_clean_img(u, p)
print("general_l3 (이창호 9단): Ready")
