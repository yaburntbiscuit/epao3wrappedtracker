from flask import Flask, render_template, request, redirect, url_for, send_file, session, flash
import zipfile
import io
import os
import sys
import requests
from bs4 import BeautifulSoup
import pandas as pd
import ast
from collections import Counter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib.pagesizes import A4
from PIL import Image, ImageDraw, ImageFont
from functools import wraps
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload
import io
import time

# Flask setup
app = Flask(__name__)
app.secret_key = 'your-secret-key-change-this-later'  # We'll make this more secure when we add password protection

# Configuration
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, 'data')
STATIC_DIR = os.path.join(BASE_DIR, 'static')
OUTPUT_DIR = os.path.join(BASE_DIR, 'output')

# Create necessary directories
os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# File paths
TRACKER_FILE_PATH = os.path.join(DATA_DIR, 'AO3_fanfiction_tracker.csv')
FONT_PATH = os.path.join(STATIC_DIR, 'LeagueSpartan.otf')

# Google Drive setup
SCOPES = ['https://www.googleapis.com/auth/drive']

SERVICE_ACCOUNT_FILE = os.path.join(
    BASE_DIR,
    "google_credentials.json"
)

credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE,
    scopes=SCOPES
)

drive_service = build('drive','v3',credentials=credentials)

# Sync settings
SYNC_INTERVAL = 300  # 5 minutes in seconds
IS_RENDER = os.environ.get('RENDER') is not None

def download_tracker_from_drive():
    """Download tracker from Google Drive with timeout and error handling"""
    try:
        file_id = "139y1kGDj7mDK6PAS77i2SOjyVM6ZkszJ"
        
        request = drive_service.files().get_media(fileId=file_id)
        fh = io.BytesIO()
        
        # Add timeout to the downloader
        downloader = MediaIoBaseDownload(fh, request, chunksize=1024*1024)  # 1MB chunks
        
        done = False
        while not done:
            status, done = downloader.next_chunk()
            if status:
                print(f"Download progress: {int(status.progress() * 100)}%")
        
        with open(TRACKER_FILE_PATH, 'wb') as f:
            f.write(fh.getvalue())
            
        print(f"Successfully downloaded tracker from Google Drive ({len(fh.getvalue())} bytes)")
        
    except Exception as e:
        print(f"Error downloading from Google Drive: {e}")
        raise

def upload_tracker_to_drive(file_path=None):
    """Upload tracker to Google Drive with error handling"""
    if file_path is None:
        file_path = TRACKER_FILE_PATH
    try:
        file_id = "139y1kGDj7mDK6PAS77i2SOjyVM6ZkszJ"
        
        media = MediaFileUpload(
            file_path,
            mimetype='text/csv',
            resumable=True  # Enable resumable uploads for reliability
        )
        
        drive_service.files().update(
            fileId=file_id,
            media_body=media
        ).execute()
        
        file_size = os.path.getsize(file_path)
        print(f"Successfully uploaded tracker to Google Drive ({file_size} bytes)")
        
    except Exception as e:
        print(f"Error uploading to Google Drive: {e}")
        print(f"File path: {file_path}")
        print(f"File exists: {os.path.exists(file_path)}")
        if os.path.exists(file_path):
            print(f"File size: {os.path.getsize(file_path)}")
        raise

# Font setup
font_144 = ImageFont.truetype(FONT_PATH, 144)
font_120 = ImageFont.truetype(FONT_PATH, 120)
font_80 = ImageFont.truetype(FONT_PATH, 80)
font_72 = ImageFont.truetype(FONT_PATH, 72)
font_40 = ImageFont.truetype(FONT_PATH, 40)
font_37 = ImageFont.truetype(FONT_PATH, 37)
font_35 = ImageFont.truetype(FONT_PATH, 35)
font_32 = ImageFont.truetype(FONT_PATH, 32)

# List of columns that should contain lists
LIST_COLUMNS = ['authors', 'ratings', 'archive_warnings', 'category', 'fandoms', 'relationships', 'characters', 'free_form_tags']


def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not session.get('logged_in'):
            return render_template('login.html')
        return f(*args, **kwargs)
    return decorated_function


# ============================================
# PASTE YOUR COLAB FUNCTIONS BELOW THIS LINE
# ============================================

# TODO: Paste these functions from your Colab code:
# - hex_to_rgb()
# - draw_centered_text()
# - draw_wrapped_centered_text()
# - extract_fic_data()

def hex_to_rgb(hex_color):
    hex_color = hex_color.lstrip("#")
    return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))


def draw_centered_text(draw, box, text, font, fill):
    x1, y1, x2, y2 = box
    bbox = draw.textbbox((0, 0), text, font=font)
    w = bbox[2] - bbox[0]
    h = bbox[3] - bbox[1]
    x = x1 + (x2 - x1 - w) // 2
    y = y1 + (y2 - y1 - h) // 2
    draw.text((x, y), text, font=font, fill=fill)


def draw_wrapped_centered_text(draw, box, text, font, fill, line_spacing=10):
    """
    Draws text wrapped and centered inside a bounding box.
    box = (x1, y1, x2, y2)
    """
    x1, y1, x2, y2 = box
    max_width = x2 - x1
    max_height = y2 - y1

    words = text.split()
    lines = []
    current_line = ""

    # Build lines word by word
    for word in words:
        test_line = current_line + (" " if current_line else "") + word
        bbox = draw.textbbox((0, 0), test_line, font=font)
        line_width = bbox[2] - bbox[0]

        if line_width <= max_width:
            current_line = test_line
        else:
            if current_line:
                lines.append(current_line)
            current_line = word

    if current_line:
        lines.append(current_line)

    # Measure total height of all lines
    line_heights = []
    total_text_height = 0

    for line in lines:
        bbox = draw.textbbox((0, 0), line, font=font)
        h = bbox[3] - bbox[1]
        line_heights.append(h)
        total_text_height += h

    total_text_height += line_spacing * (len(lines) - 1)

    # Starting Y so the block is vertically centered
    y = y1 + (max_height - total_text_height) // 2

    # Draw each line centered horizontally
    for line, h in zip(lines, line_heights):
        bbox = draw.textbbox((0, 0), line, font=font)
        w = bbox[2] - bbox[0]
        x = x1 + (max_width - w) // 2
        draw.text((x, y), line, font=font, fill=fill)
        y += h + line_spacing


def extract_fic_data(url):
    """
    Extracts detailed information from an AO3 fanfiction URL.

    Args:
        url (str): The URL of the AO3 fanfiction.

    Returns:
        dict: A dictionary containing extracted fanfiction data.
              Returns an empty dictionary if extraction fails.
    """
    fic_data = {
        'url': url,
        'title': None, # Added for title extraction
        'word_count': None,
        'ratings': [],
        'archive_warnings': [],
        'category': [],
        'fandoms': [],
        'relationships': [],
        'characters': [],
        'free_form_tags': [],
        'authors': [] # Added authors key
    }

    try:
        response = requests.get(url)
        response.raise_for_status()  # Raise an HTTPError for bad responses (4xx or 5xx)
        soup = BeautifulSoup(response.text, 'html.parser')

        # Extract Title
        title_h2 = soup.find('h2', class_='title')
        if title_h2:
            fic_data['title'] = title_h2.text.strip()

        # Find the work meta group where most of the data is located
        work_meta = soup.find('dl', class_='work meta group')

        if work_meta:
            # Extract Word Count
            word_count_dd = work_meta.find('dd', class_='words')
            if word_count_dd:
                fic_data['word_count'] = int(word_count_dd.text.replace(',', ''))

            # Extract Ratings
            rating_dd = work_meta.find('dd', class_='rating tags')
            if rating_dd:
                fic_data['ratings'] = [a.text for a in rating_dd.find_all('a')]

            # Extract Archive Warnings
            warning_dd = work_meta.find('dd', class_='warning tags')
            if warning_dd:
                fic_data['archive_warnings'] = [a.text for a in warning_dd.find_all('a')]

            # Extract Category
            category_dd = work_meta.find('dd', class_='category tags')
            if category_dd:
                fic_data['category'] = [a.text for a in category_dd.find_all('a')]

            # Extract Fandoms
            fandom_dd = work_meta.find('dd', class_='fandom tags')
            if fandom_dd:
                fic_data['fandoms'] = [a.text for a in fandom_dd.find_all('a')]

            # Extract Relationships
            relationship_dd = work_meta.find('dd', class_='relationship tags')
            if relationship_dd:
                fic_data['relationships'] = [a.text for a in relationship_dd.find_all('a')]

            # Extract Characters
            character_dd = work_meta.find('dd', class_='character tags')
            if character_dd:
                fic_data['characters'] = [a.text for a in character_dd.find_all('a')]

            # Extract Freeform Tags
            freeform_dd = work_meta.find('dd', class_='freeform tags')
            if freeform_dd:
                fic_data['free_form_tags'] = [a.text for a in freeform_dd.find_all('a')]

        # Extract Author(s)
        # Changed from soup.find('div', class_='byline') to soup.find(class_='byline') for broader matching
        byline_element = soup.find(class_='byline')
        if byline_element:
            author_links = byline_element.find_all('a', rel='author')
            fic_data['authors'] = [a.text for a in author_links]

    except requests.exceptions.RequestException as e:
        print(f"Request failed for {url}: {e}")
    except Exception as e:
        print(f"An error occurred during extraction for {url}: {e}")

    return fic_data


# ============================================
# HELPER FUNCTIONS FOR WEB APP
# ============================================

# Global variable to track if we need to sync with Drive
LAST_SYNC_TIME = 0
SYNC_INTERVAL = 300  # 5 minutes in seconds

def load_tracker():
    """Load the tracker CSV file with smart caching"""
    if IS_RENDER:
        # On Render, always download from Drive since local files don't persist
        try:
            download_tracker_from_drive()
        except Exception as e:
            print(f"Error: Could not download from Google Drive: {e}")
            # Return empty DataFrame if Drive fails
            return pd.DataFrame(columns=[
                'url', 'title', 'word_count', 'authors', 'ratings', 'archive_warnings', 'category',
                'fandoms', 'relationships', 'characters', 'free_form_tags'
            ])
        
        # Load the freshly downloaded file
        df = pd.read_csv(TRACKER_FILE_PATH)
        # Convert string representations of lists back to actual lists
        for col in LIST_COLUMNS:
            if col in df.columns:
                df[col] = df[col].apply(lambda x: ast.literal_eval(x) if isinstance(x, str) else x)
        return df
    
    # Local development: use caching
    current_time = time.time()
    
    # Check if local file exists and is recent enough (within last 5 minutes)
    if os.path.exists(TRACKER_FILE_PATH):
        file_mod_time = os.path.getmtime(TRACKER_FILE_PATH)
        if (current_time - file_mod_time) < SYNC_INTERVAL:
            # File is recent, load from local
            df = pd.read_csv(TRACKER_FILE_PATH)
            # Convert string representations of lists back to actual lists
            for col in LIST_COLUMNS:
                if col in df.columns:
                    df[col] = df[col].apply(lambda x: ast.literal_eval(x) if isinstance(x, str) else x)
            return df
    
    # File doesn't exist or is old, download from Drive
    try:
        download_tracker_from_drive()
    except Exception as e:
        print(f"Warning: Could not sync with Google Drive: {e}")
        # If Drive fails but local file exists, use local file
        if os.path.exists(TRACKER_FILE_PATH):
            df = pd.read_csv(TRACKER_FILE_PATH)
            # Convert string representations of lists back to actual lists
            for col in LIST_COLUMNS:
                if col in df.columns:
                    df[col] = df[col].apply(lambda x: ast.literal_eval(x) if isinstance(x, str) else x)
            return df
        # If no local file either, create empty DataFrame
        return pd.DataFrame(columns=[
            'url', 'title', 'word_count', 'authors', 'ratings', 'archive_warnings', 'category',
            'fandoms', 'relationships', 'characters', 'free_form_tags'
        ])
    
    # Load the freshly downloaded file
    df = pd.read_csv(TRACKER_FILE_PATH)
    # Convert string representations of lists back to actual lists
    for col in LIST_COLUMNS:
        if col in df.columns:
            df[col] = df[col].apply(lambda x: ast.literal_eval(x) if isinstance(x, str) else x)
    return df


def save_tracker(df):
    """Save the tracker DataFrame to CSV and sync with Google Drive"""
    global LAST_SYNC_TIME
    df_to_save = df.copy()
    for col in LIST_COLUMNS:
        if col in df_to_save.columns:
            df_to_save[col] = df_to_save[col].apply(lambda x: str(x) if isinstance(x, list) else x)
    
    if not IS_RENDER:
        df_to_save.to_csv(TRACKER_FILE_PATH, index=False)
        # On local, try to upload but don't fail if it doesn't work
        try:
            upload_tracker_to_drive()
            LAST_SYNC_TIME = time.time()  # Update sync time after successful upload
        except Exception as e:
            print(f"Warning: Failed to sync with Google Drive, but local save succeeded: {e}")
    else:
        # On Render, save temporarily to upload, then delete
        temp_path = TRACKER_FILE_PATH + '.temp'
        df_to_save.to_csv(temp_path, index=False)
        try:
            upload_tracker_to_drive(temp_path)
            LAST_SYNC_TIME = time.time()  # Update sync time after successful upload
        finally:
            if os.path.exists(temp_path):
                os.remove(temp_path)


def add_fic_to_tracker(url):
    """Add a fic to the tracker"""
    df = load_tracker()
    
    print(f"Extracting data for: {url}")
    fic_details = extract_fic_data(url)
    
    if not fic_details or fic_details.get('word_count') is None:
        return False, f"❌ Could not extract data for {url}"
    
    new_fic_series = pd.Series(fic_details).reindex(df.columns)
    df = pd.concat([df, pd.DataFrame([new_fic_series])], ignore_index=True)
    try:
        save_tracker(df)
    except Exception as e:
        return False, f"❌ Failed to save fic: {e}"
    
    return True, f"✅ Successfully added: {fic_details.get('title', 'Unknown')}"


def add_fic_manually(fic_data):
    """Add a fic to the tracker manually (from form submission)"""
    df = load_tracker()
    
    # Create series with provided data
    new_fic_series = pd.Series(fic_data).reindex(df.columns)
    df = pd.concat([df, pd.DataFrame([new_fic_series])], ignore_index=True)
    try:
        save_tracker(df)
    except Exception as e:
        return False, f"❌ Failed to save fic: {e}"
    
    return True, f"✅ Successfully added: {fic_data.get('title', 'Unknown')}"


# ============================================
# WEB ROUTES
# ============================================

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        password = request.form.get('password')
        if password == 'HaydenIsTheAlpha123!':  # Change this to your desired password
            session.permanent = False  # Session expires when browser closes
            session['logged_in'] = True
            return redirect(url_for('index'))
        else:
            flash('Invalid password')
            return render_template('login.html')
    return render_template('login.html')


@app.route('/logout')
def logout():
    """Log out the user"""
    session.clear()
    return redirect(url_for('login'))


@app.route('/')
@login_required
def index():
    """Main page"""
    df = load_tracker()
    total_fics = len(df)
    return render_template('index1.html', total_fics=total_fics)


@app.route('/add_fic', methods=['POST'])
@login_required
def add_fic():
    """Handle adding fics from the form"""
    urls_input = request.form.get('urls', '')
    urls = [url.strip() for url in urls_input.split('\n') if url.strip()]
    
    messages = []
    for url in urls:
        success, message = add_fic_to_tracker(url)
        messages.append(message)
    
    return render_template('results.html', action='add', results=messages, show_downloads=False)


@app.route('/generate_report', methods=['POST'])
@login_required
def generate_report():
    """Generate the wrapped report"""
    df = load_tracker()
    
    if df.empty:
        return render_template('results.html', action='generate', results=["No fics to generate report from."], show_downloads=False)
    
    # Compile data
    total_word_count = df['word_count'].sum()
    total_unique_fics = df['url'].nunique()
    average_words_per_day = round(total_word_count / 365)
    
    # Longest fic
    longest_fic = df.loc[df['word_count'].idxmax()]
    longest_fic_title = longest_fic['title']
    longest_fic_word_count = longest_fic['word_count']
    longest_fic_url = longest_fic['url']
    longest_fic_author = longest_fic['authors'][0] if longest_fic['authors'] else 'N/A'
    
    # Favorite fic
    url_counts = df['url'].value_counts()
    most_frequent_url = url_counts.idxmax()
    favorite_fic_visits = url_counts.max()
    favorite_fic = df[df['url'] == most_frequent_url].iloc[0]
    favorite_fic_title = favorite_fic['title']
    favorite_fic_word_count = favorite_fic['word_count']
    favorite_fic_author = favorite_fic['authors'][0] if favorite_fic['authors'] else 'N/A'
    favorite_fic_url = favorite_fic['url']
    
    # Top 5 authors
    all_authors = []
    for authors_list in df['authors']:
        all_authors.extend(authors_list)
    author_counts = Counter(all_authors)
    top_5_authors = [author for author, _ in author_counts.most_common(5)]
    
    # Top 5 fandoms
    all_fandoms = []
    for fandoms_list in df['fandoms']:
        all_fandoms.extend(fandoms_list)
    fandom_counts = Counter(all_fandoms)
    top_5_fandoms = [fandom for fandom, _ in fandom_counts.most_common(5)]
    
    # Top 5 relationships
    all_relationships = []
    for relationships_list in df['relationships']:
        all_relationships.extend(relationships_list)
    relationship_counts = Counter(all_relationships)
    top_5_relationships = [rel for rel, _ in relationship_counts.most_common(5)]
    
    # Top 10 tags
    all_free_form_tags = []
    for tags_list in df['free_form_tags']:
        all_free_form_tags.extend(tags_list)
    tag_counts = Counter(all_free_form_tags)
    top_10_tags = [tag for tag, _ in tag_counts.most_common(10)]
    
    # Generate PDF
    pdf_filename = os.path.join(OUTPUT_DIR, 'AO3_Wrapped_Report.pdf')
    doc = SimpleDocTemplate(pdf_filename, pagesize=A4)
    story = []
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name='ReportHeading1', fontSize=24, leading=28, alignment=1, spaceAfter=20, fontName='Helvetica-Bold'))
    styles.add(ParagraphStyle(name='ReportHeading2', fontSize=18, leading=22, spaceAfter=10, fontName='Helvetica-Bold'))
    styles.add(ParagraphStyle(name='BodyTextIndent', parent=styles['Normal'], firstLineIndent=0.25*inch, spaceAfter=6))
    styles.add(ParagraphStyle(name='ListBullet', parent=styles['Normal'], leftIndent=0.5*inch, spaceAfter=3))
    
    story.append(Paragraph('AO3 Wrapped Report 2026', styles['ReportHeading1']))
    story.append(Spacer(1, 0.2 * inch))
    story.append(Paragraph('Overall Statistics', styles['ReportHeading2']))
    story.append(Paragraph(f"Total Word Count: {total_word_count:,}", styles['Normal']))
    story.append(Paragraph(f"Total Unique Fics: {total_unique_fics}", styles['Normal']))
    story.append(Paragraph(f"Average Words Per Day (based on 365 days): {average_words_per_day}", styles['Normal']))
    story.append(Spacer(1, 0.2 * inch))
    
    story.append(Paragraph('Longest Fanfiction', styles['ReportHeading2']))
    story.append(Paragraph(f"<b>Title:</b> {longest_fic_title}", styles['Normal']))
    story.append(Paragraph(f"<b>Word Count:</b> {longest_fic_word_count:,}", styles['Normal']))
    story.append(Paragraph(f"<b>Author:</b> {longest_fic_author}", styles['Normal']))
    story.append(Paragraph(f"<b>URL:</b> {longest_fic_url}", styles['Normal']))
    story.append(Spacer(1, 0.2 * inch))
    
    story.append(Paragraph('Favorite Fanfiction', styles['ReportHeading2']))
    story.append(Paragraph(f"<b>Title:</b> {favorite_fic_title}", styles['Normal']))
    story.append(Paragraph(f"<b>Word Count:</b> {favorite_fic_word_count:,}", styles['Normal']))
    story.append(Paragraph(f"<b>Author:</b> {favorite_fic_author}", styles['Normal']))
    story.append(Paragraph(f"<b>URL:</b> {favorite_fic_url}", styles['Normal']))
    story.append(Paragraph(f"<b>Visits:</b> {favorite_fic_visits}", styles['Normal']))
    story.append(Spacer(1, 0.2 * inch))
    
    story.append(Paragraph('Top 5 Authors', styles['ReportHeading2']))
    for i, author in enumerate(top_5_authors, 1):
        story.append(Paragraph(f"{i}. {author}", styles['Normal']))
    story.append(Spacer(1, 0.2 * inch))
    
    story.append(Paragraph('Top 5 Fandoms', styles['ReportHeading2']))
    for i, fandom in enumerate(top_5_fandoms, 1):
        story.append(Paragraph(f"{i}. {fandom}", styles['Normal']))
    story.append(Spacer(1, 0.2 * inch))
    
    story.append(Paragraph('Top 5 Relationships', styles['ReportHeading2']))
    for i, rel in enumerate(top_5_relationships, 1):
        story.append(Paragraph(f"{i}. {rel}", styles['Normal']))
    story.append(Spacer(1, 0.2 * inch))
    
    story.append(Paragraph('Top 10 Freeform Tags', styles['ReportHeading2']))
    for i, tag in enumerate(top_10_tags, 1):
        story.append(Paragraph(f"{i}. {tag}", styles['Normal']))
    story.append(Spacer(1, 0.2 * inch))
    
    doc.build(story)
    
    # Generate images
    # Page 1
    page1_template = os.path.join(STATIC_DIR, '2026_AO3_Wrapped_Template_Page1.png')
    img = Image.open(page1_template)
    img.save(os.path.join(OUTPUT_DIR, 'wrapped_page1.png'))
    
    # Page 2
    page2_template = os.path.join(STATIC_DIR, '2026_AO3_Wrapped_Template_Page2.png')
    img = Image.open(page2_template)
    draw = ImageDraw.Draw(img)
    draw.text((30, 339), f"{total_word_count:,}", font=font_144, fill=hex_to_rgb("#f7f7f7"))
    draw.text((30, 660), f"{total_unique_fics}", font=font_120, fill=hex_to_rgb("#f7f7f7"))
    draw.text((30, 1090), f"{average_words_per_day}", font=font_80, fill=hex_to_rgb("#f7f7f7"))
    img.save(os.path.join(OUTPUT_DIR, 'wrapped_page2.png'))
    
    # Page 3
    page3_template = os.path.join(STATIC_DIR, '2026_AO3_Wrapped_Template_Page3.png')
    img = Image.open(page3_template)
    draw = ImageDraw.Draw(img)
    fav_title_text = favorite_fic_title
    fav_author_text = f"by {favorite_fic_author}"
    fav_visits_text = f"You visited this fic {favorite_fic_visits} times!"
    draw_wrapped_centered_text(draw, (90, 406, 988, 895), fav_title_text, font_72, hex_to_rgb("#333333"), 12)
    draw_centered_text(draw, (216, 920, 863, 1066), fav_author_text, font_40, hex_to_rgb("#333333"))
    draw_centered_text(draw, (35, 1218, 1045, 1318), fav_visits_text, font_37, hex_to_rgb("#333333"))
    img.save(os.path.join(OUTPUT_DIR, 'wrapped_page3.png'))
    
    # Page 4
    page4_template = os.path.join(STATIC_DIR, '2026_AO3_Wrapped_Template_Page4.png')
    img = Image.open(page4_template)
    draw = ImageDraw.Draw(img)
    top_5_authors_text = "\n".join(f"{i+1}. {author}" for i, author in enumerate(top_5_authors))
    draw.text((180, 282), top_5_authors_text, font=font_35, fill=hex_to_rgb("#f7f7f7"), spacing=12)
    img.save(os.path.join(OUTPUT_DIR, 'wrapped_page4.png'))
    
    # Page 5
    page5_template = os.path.join(STATIC_DIR, '2026_AO3_Wrapped_Template_Page5.png')
    img = Image.open(page5_template)
    draw = ImageDraw.Draw(img)
    top_5_relationships_text = "\n".join(f"{i+1}. {fandom}" for i, fandom in enumerate(top_5_relationships))
    draw.text((128, 374), top_5_relationships_text, font=font_35, fill=hex_to_rgb("#f7f7f7"), spacing=12)
    img.save(os.path.join(OUTPUT_DIR, 'wrapped_page5.png'))

    # Page 6
    page6_template = os.path.join(STATIC_DIR, '2026_AO3_Wrapped_Template_Page6.png')
    img = Image.open(page6_template)
    draw = ImageDraw.Draw(img)
    tags_1_to_5 = top_10_tags[:5]
    tags_6_to_10 = top_10_tags[5:10]
    tags_1_to_5_text = "\n".join(f"{i+1}. {tag}" for i, tag in enumerate(tags_1_to_5))
    tags_6_to_10_text = "\n".join(f"{i+6}. {tag}" for i, tag in enumerate(tags_6_to_10))
    draw.text((95, 190), tags_1_to_5_text, font=font_35, fill=hex_to_rgb("#333333"), spacing=20)
    draw.text((84, 815), tags_6_to_10_text, font=font_35, fill=hex_to_rgb("#333333"), spacing=20)
    img.save(os.path.join(OUTPUT_DIR, 'wrapped_page6.png'))

    return render_template('results.html', action='generate', results=[], show_downloads=True)


@app.route('/download/pdf')
def download_pdf():
    pdf_path = os.path.join(OUTPUT_DIR, 'AO3_Wrapped_Report.pdf')
    if os.path.exists(pdf_path):
        return send_file(pdf_path, as_attachment=True)
    else:
        return "PDF not found", 404

@app.route('/download/images')
def download_images():
    images = ['wrapped_page1.png', 'wrapped_page2.png', 'wrapped_page3.png', 'wrapped_page4.png', 'wrapped_page5.png']
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w') as zip_file:
        for img in images:
            img_path = os.path.join(OUTPUT_DIR, img)
            if os.path.exists(img_path):
                zip_file.write(img_path, img)
    zip_buffer.seek(0)
    return send_file(zip_buffer, mimetype='application/zip', as_attachment=True, download_name='wrapped_images.zip')


@app.route('/download/<filename>')
@login_required
def download_file(filename):
    file_path = os.path.join(OUTPUT_DIR, filename)
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    else:
        return "File not found", 404


@app.route('/back_to_tracker')
@login_required
def back_to_tracker():
    """Clean up generated files and redirect to tracker"""
    import glob
    
    # Delete PDF
    pdf_pattern = os.path.join(OUTPUT_DIR, 'AO3_Wrapped_Report.pdf')
    for f in glob.glob(pdf_pattern):
        try:
            os.remove(f)
        except OSError:
            pass  # Ignore if file doesn't exist or can't be deleted
    
    # Delete PNG files
    png_pattern = os.path.join(OUTPUT_DIR, 'wrapped_page*.png')
    for f in glob.glob(png_pattern):
        try:
            os.remove(f)
        except OSError:
            pass
    
    return redirect(url_for('index'))


@app.route('/manual_entry')
@login_required
def manual_entry():
    """Display manual entry form"""
    return render_template('manual_entry.html')


@app.route('/add_fic_manual', methods=['POST'])
@login_required
def add_fic_manual():
    """Handle manual fic entry from form"""
    try:
        # Required fields
        url = request.form.get('url', '').strip()
        title = request.form.get('title', '').strip()
        word_count_str = request.form.get('word_count', '').strip()
        authors_str = request.form.get('authors', '').strip()
        
        # Validate required fields
        if not url or not title or not word_count_str or not authors_str:
            flash('Please fill out all required fields (URL, Title, Word Count, Authors)')
            return render_template('manual_entry.html')
        
        # Convert word count to integer
        try:
            word_count = int(word_count_str)
        except ValueError:
            flash('Word count must be a valid number')
            return render_template('manual_entry.html')
        
        # Optional fields
        ratings = request.form.getlist('ratings')
        archive_warnings = request.form.getlist('archive_warnings')
        category = request.form.getlist('category')
        
        # Parse comma-separated fields
        fandoms = [f.strip() for f in request.form.get('fandoms', '').split(',') if f.strip()]
        relationships = [r.strip() for r in request.form.get('relationships', '').split(',') if r.strip()]
        characters = [c.strip() for c in request.form.get('characters', '').split(',') if c.strip()]
        free_form_tags = [t.strip() for t in request.form.get('free_form_tags', '').split(',') if t.strip()]
        authors = [a.strip() for a in authors_str.split(',') if a.strip()]
        
        # Create fic data dictionary
        fic_data = {
            'url': url,
            'title': title,
            'word_count': word_count,
            'ratings': ratings,
            'archive_warnings': archive_warnings,
            'category': category,
            'fandoms': fandoms,
            'relationships': relationships,
            'characters': characters,
            'free_form_tags': free_form_tags,
            'authors': authors
        }
        
        # Add to tracker
        success, message = add_fic_manually(fic_data)
        
        # Redirect to results page with success message
        return render_template('results.html', action='add', results=[message], show_downloads=False)
        
    except Exception as e:
        print(f"Error in manual entry: {e}")
        flash(f'An error occurred: {str(e)}')
        return render_template('manual_entry.html')


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)