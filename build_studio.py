# -*- coding: utf-8 -*-
import json
import re
import os
import sys
from pathlib import Path

# Set working directory to project root
PROJECT_DIR = Path(r"c:\Users\magro\Downloads\Fitness Programmes")
os.chdir(PROJECT_DIR)
sys.path.insert(0, str(PROJECT_DIR))

from main import build_app_data, derive_client_name, generate_complete_html

# 1. Extract verified CSS & JS from Jada Said.html
if not Path('Jada Said.html').exists():
    generate_complete_html('Jada Said.xlsx')

with open('Jada Said.html', 'r', encoding='utf-8') as f:
    client_html = f.read()

client_css = re.search(r'<style>(.*?)</style>', client_html, re.DOTALL).group(1)
client_js = re.search(r'<script>(.*?)</script>', client_html, re.DOTALL).group(1)

# Preload Excel data for Jada Said
excel_file = 'Jada Said.xlsx'
jada_pages = build_app_data(excel_file)

initial_athletes = [
    {
        'id': 'ath_1',
        'name': 'Jada Said',
        'initials': 'JS',
        'comments': 'Demonstrating excellent scapular control on rows. Working on breathing and core bracing control on compound lifts.',
        'pages': jada_pages,
        'logs': [
            {
                'id': 'log-js-1',
                'athleteId': 'ath_1',
                'athleteName': 'Jada Said',
                'athleteInitials': 'JS',
                'sessionId': 'page-1',
                'sessionTitle': 'Session 1 - Upper Power',
                'date': '2026-09-02',
                'status': 'completed',
                'exercisesSummary': '4 exercises  -  12 sets',
                'exercises': [
                    {'exercise': 'Bent-Elbow Band Pull-Aparts', 'sets': '2 sets (Light Band x 10, Light Band x 10)'},
                    {'exercise': 'DB Bench Press', 'sets': '4 sets (14kg x 10, 14kg x 10, 14kg x 10, 14kg x 8)'},
                    {'exercise': 'Lat Pulldown', 'sets': '3 sets (35kg x 10, 35kg x 10, 35kg x 10)'},
                    {'exercise': 'Seated Cable Row', 'sets': '3 sets (30kg x 12, 30kg x 12, 30kg x 10)'}
                ],
                'notes': 'Solid upper body power session. Felt strong on DB Bench Press.',
                'loggedAt': '2026-09-02T10:30:00.000Z'
            },
            {
                'id': 'log-js-2',
                'athleteId': 'ath_1',
                'athleteName': 'Jada Said',
                'athleteInitials': 'JS',
                'sessionId': 'page-2',
                'sessionTitle': 'Session 2 - Lower Strength',
                'date': '2026-09-04',
                'status': 'completed',
                'exercisesSummary': '3 exercises  -  10 sets',
                'exercises': [
                    {'exercise': 'Goblet Squat', 'sets': '4 sets (20kg x 8, 20kg x 8, 24kg x 8, 24kg x 6)'},
                    {'exercise': 'Romanian Deadlift', 'sets': '3 sets (40kg x 10, 45kg x 10, 45kg x 8)'},
                    {'exercise': 'Walking Lunges', 'sets': '3 sets (12kg x 10, 12kg x 10, 12kg x 10)'}
                ],
                'notes': 'Great depth on squats. Increased RDL load.',
                'loggedAt': '2026-09-04T11:00:00.000Z'
            },
            {
                'id': 'log-js-3',
                'athleteId': 'ath_1',
                'athleteName': 'Jada Said',
                'athleteInitials': 'JS',
                'sessionId': 'page-1',
                'sessionTitle': 'Session 1 - Upper Power',
                'date': '2026-09-07',
                'status': 'scheduled',
                'exercisesSummary': '4 exercises  -  12 sets',
                'exercises': [],
                'notes': 'Upcoming scheduled upper session.',
                'loggedAt': '2026-09-06T09:00:00.000Z'
            }
        ]
    },
    {
        'id': 'ath_2',
        'name': 'Marcus Vance',
        'initials': 'MV',
        'comments': 'Focus on deceleration mechanics and knee stability during plyometric landings.',
        'pages': [
            {
                'id': 'page-mv-1',
                'title': 'Session 1 - Lower Power',
                'nav_title': 'S 1 - Lower',
                'type': 'exercises',
                'cards': [
                    {
                        'category': 'Warm-Up',
                        'exercise': "World's Greatest Stretch",
                        'details': '2 x 5 per side',
                        'weight': 'Bodyweight',
                        'rest': '30s',
                        'rest_seconds': 30,
                        'total_sets': 2,
                        'cue': 'Open thoracic spine tall toward ceiling.',
                        'videos': []
                    },
                    {
                        'category': 'Primary Strength',
                        'exercise': 'Barbell Back Squat',
                        'details': '4 x 6-8',
                        'weight': '110kg to 125kg',
                        'rest': '120s',
                        'rest_seconds': 120,
                        'total_sets': 4,
                        'cue': 'Explode out of the hole driving through midfoot.',
                        'videos': []
                    }
                ]
            },
            {
                'id': 'page-mv-2',
                'title': 'Session 2 - Upper Strength',
                'nav_title': 'S 2 - Upper',
                'type': 'exercises',
                'cards': [
                    {
                        'category': 'Primary Strength',
                        'exercise': 'DB Bent-Over Row',
                        'details': '4 x 8-10',
                        'weight': '22kg',
                        'rest': '90s',
                        'rest_seconds': 90,
                        'total_sets': 4,
                        'cue': 'Pull to hip pocket with flat spine.',
                        'videos': []
                    }
                ]
            }
        ],
        'logs': [
            {
                'id': 'log-mv-1',
                'athleteId': 'ath_2',
                'athleteName': 'Marcus Vance',
                'athleteInitials': 'MV',
                'sessionId': 'page-mv-1',
                'sessionTitle': 'Session 1 - Lower Power',
                'date': '2026-09-03',
                'status': 'completed',
                'exercisesSummary': '2 exercises  -  6 sets',
                'exercises': [
                    {'exercise': "World's Greatest Stretch", 'sets': '2 sets (Bodyweight x 5, Bodyweight x 5)'},
                    {'exercise': 'Barbell Back Squat', 'sets': '4 sets (110kg x 8, 115kg x 8, 120kg x 6, 125kg x 6)'}
                ],
                'notes': 'Explosive squatting with smooth depth.',
                'loggedAt': '2026-09-03T16:00:00.000Z'
            },
            {
                'id': 'log-mv-2',
                'athleteId': 'ath_2',
                'athleteName': 'Marcus Vance',
                'athleteInitials': 'MV',
                'sessionId': 'page-mv-2',
                'sessionTitle': 'Session 2 - Upper Strength',
                'date': '2026-09-05',
                'status': 'completed',
                'exercisesSummary': '1 exercise  -  4 sets',
                'exercises': [
                    {'exercise': 'DB Bent-Over Row', 'sets': '4 sets (22kg x 10, 22kg x 10, 22kg x 8, 22kg x 8)'}
                ],
                'notes': 'Solid back engagement and control.',
                'loggedAt': '2026-09-05T15:30:00.000Z'
            }
        ]
    }
]

initial_athletes_json = json.dumps(initial_athletes, indent=2)
client_css_json = json.dumps(client_css)
client_js_json = json.dumps(client_js)

MUI_ICONS = {
    'bolt': '<svg class="mui-icon" viewBox="0 0 24 24"><path d="M11 21h-1l1-7H7.5c-.58 0-.57-.32-.38-.66.19-.34.05-.08.07-.12C8.48 10.94 10.42 7.54 13 3h1l-1 7h3.5c.49 0 .56.33.47.51l-.07.15C12.9 17.5 11 21 11 21z"/></svg>',
    'athletes': '<svg class="mui-icon" viewBox="0 0 24 24"><path d="M16 11c1.66 0 2.99-1.34 2.99-3S17.66 5 16 5c-1.66 0-3 1.34-3 3s1.34 3 3 3zm-8 0c1.66 0 2.99-1.34 2.99-3S9.66 5 8 5C6.34 5 5 6.34 5 8s1.34 3 3 3zm0 2c-2.33 0-7 1.17-7 3.5V19h14v-2.5c0-2.33-4.67-3.5-7-3.5zm8 0c-.29 0-.62.02-.97.05 1.16.84 1.97 1.97 1.97 3.45V19h6v-2.5c0-2.33-4.67-3.5-7-3.5z"/></svg>',
    'templates': '<svg class="mui-icon" viewBox="0 0 24 24"><path d="M19 3h-4.18C14.4 1.84 13.3 1 12 1c-1.3 0-2.4.84-2.82 2H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm-7 0c.55 0 1 .45 1 1s-.45 1-1 1-1-.45-1-1 .45-1 1-1zm2 14H7v-2h7v2zm3-4H7v-2h10v2zm0-4H7V7h10v2z"/></svg>',
    'library': '<svg class="mui-icon" viewBox="0 0 24 24"><path d="M4 6H2v14c0 1.1.9 2 2 2h14v-2H4V6zm16-4H8c-1.1 0-2 .9-2 2v12c0 1.1.9 2 2 2h12c1.1 0 2-.9 2-2V4c0-1.1-.9-2-2-2zm0 14H8V4h12v12z"/></svg>',
    'add': '<svg class="mui-icon" viewBox="0 0 24 24"><path d="M19 13h-6v6h-2v-6H5v-2h6V5h2v6h6v2z"/></svg>',
    'search': '<svg class="mui-icon" viewBox="0 0 24 24"><path d="M15.5 14h-.79l-.28-.27A6.471 6.471 0 0 0 16 9.5 6.5 6.5 0 1 0 9.5 16c1.61 0 3.09-.59 4.23-1.57l.27.28v.79l5 4.99L20.49 19l-4.99-5zm-6 0C7.01 14 5 11.99 5 9.5S7.01 5 9.5 5 14 7.01 14 9.5 11.99 14 9.5 14z"/></svg>',
    'edit': '<svg class="mui-icon" viewBox="0 0 24 24"><path d="M3 17.25V21h3.75L17.81 9.94l-3.75-3.75L3 17.25zM20.71 7.04c.39-.39.39-1.02 0-1.41l-2.34-2.34c-.39-.39-1.02-.39-1.41 0l-1.83 1.83 3.75 3.75 1.83-1.83z"/></svg>',
    'copy': '<svg class="mui-icon" viewBox="0 0 24 24"><path d="M16 1H4c-1.1 0-2 .9-2 2v14h2V3h12V1zm3 4H8c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h11c1.1 0 2-.9 2-2V7c0-1.1-.9-2-2-2zm0 16H8V7h11v14z"/></svg>',
    'duplicate': '<svg class="mui-icon" viewBox="0 0 24 24"><path d="M15 1H4c-1.1 0-2 .9-2 2v13h2V3h11V1zm4 4H8c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h11c1.1 0 2-.9 2-2V7c0-1.1-.9-2-2-2zm0 16H8V7h11v14z"/></svg>',
    'delete': '<svg class="mui-icon" viewBox="0 0 24 24"><path d="M6 19c0 1.1.9 2 2 2h8c1.1 0 2-.9 2-2V7H6v12zM19 4h-3.5l-1-1h-5l-1 1H5v2h14V4z"/></svg>',
    'save': '<svg class="mui-icon" viewBox="0 0 24 24"><path d="M17 3H5c-1.11 0-2 .9-2 2v14c0 1.1.89 2 2 2h14c1.1 0 2-.9 2-2V7l-4-4zm-5 16c-1.66 0-3-1.34-3-3s1.34-3 3-3 3 1.34 3 3-1.34 3-3 3zm3-10H5V5h10v4z"/></svg>',
    'phone': '<svg class="mui-icon" viewBox="0 0 24 24"><path d="M17 1.01L7 1c-1.1 0-2 .9-2 2v18c0 1.1.9 2 2 2h10c1.1 0 2-.9 2-2V3c0-1.1-.9-1.99-2-1.99zM17 19H7V5h10v14z"/></svg>',
    'export': '<svg class="mui-icon" viewBox="0 0 24 24"><path d="M19 19H5V8h2v9h10V8h2v11zM12 3l5 5h-4v7h-2V8H7l5-5z"/></svg>',
    'rocket': '<svg class="mui-icon" viewBox="0 0 24 24"><path d="M19 19H5V8h2v9h10V8h2v11zM12 3l5 5h-4v7h-2V8H7l5-5z"/></svg>',
    'drag': '<svg class="mui-icon" viewBox="0 0 24 24"><path d="M11 18c0 1.1-.9 2-2 2s-2-.9-2-2 .9-2 2-2 2 .9 2 2zm-2-8c-1.1 0-2 .9-2 2s.9 2 2 2 2-.9 2-2-.9-2-2-2zm0-6c-1.1 0-2 .9-2 2s.9 2 2 2 2-.9 2-2-.9-2-2-2zm6 4c1.1 0 2-.9 2-2s-.9-2-2-2-2 .9-2 2 .9 2 2 2zm0 2c-1.1 0-2 .9-2 2s.9 2 2 2 2-.9 2-2-.9-2-2-2zm0 6c-1.1 0-2 .9-2 2s.9 2 2 2 2-.9 2-2-.9-2-2-2z"/></svg>',
    'timer': '<svg class="mui-icon" viewBox="0 0 24 24"><path d="M15 1H9v2h6V1zm-4 13h2V8h-2v6zm8.03-6.61l1.42-1.42c-.43-.51-.9-.99-1.41-1.41l-1.42 1.42C16.07 4.74 14.12 4 12 4c-4.97 0-9 4.03-9 9s4.02 9 9 9 9-4.03 9-9c0-2.12-.74-4.07-1.97-5.61zM12 20c-3.87 0-7-3.13-7-7s3.13-7 7-7 7 3.13 7 7-3.13 7-7 7z"/></svg>',
    'dumbbell': '<svg class="mui-icon" viewBox="0 0 24 24"><path d="M20.57 14.86L22 13.43 20.57 12 17 15.57 8.43 7 12 3.43 10.57 2 9.14 3.43 7.71 2 5.57 4.14 4.14 2.71 2.71 4.14l1.43 1.43L2 7.71l1.43 1.43L2 10.57 3.43 12 7 8.43 15.57 17 12 20.57 13.43 22l1.43-1.43 1.43 1.43 2.14-2.14 1.43 1.43 1.43-1.43-1.43-1.43L22 16.29l-1.43-1.43z"/></svg>',
    'trending_up': '<svg class="mui-icon" viewBox="0 0 24 24"><path d="M16 6l2.29 2.29-4.88 4.88-4-4L2 16.59 3.41 18l6-6 4 4 6.3-6.29L22 12V6z"/></svg>',
    'calendar': '<svg class="mui-icon" viewBox="0 0 24 24"><path d="M19 3h-1V1h-2v2H8V1H6v2H5c-1.11 0-1.99.9-1.99 2L3 19c0 1.1.89 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm0 16H5V8h14v11zM7 10h5v5H7z"/></svg>',
    'check': '<svg class="mui-icon" viewBox="0 0 24 24"><path d="M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41z"/></svg>',
    'warning': '<svg class="mui-icon" viewBox="0 0 24 24"><path d="M1 21h22L12 2 1 21zm12-3h-2v-2h2v2zm0-4h-2v-4h2v4z"/></svg>',
    'close': '<svg class="mui-icon" viewBox="0 0 24 24"><path d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z"/></svg>',
    'arrow_back': '<svg class="mui-icon" viewBox="0 0 24 24"><path d="M20 11H7.83l5.59-5.59L12 4l-8 8 8 8 1.41-1.41L7.83 13H20v-2z"/></svg>',
    'comment': '<svg class="mui-icon" viewBox="0 0 24 24"><path d="M21.99 4c0-1.1-.89-2-1.99-2H4c-1.1 0-2 .9-2 2v12c0 1.1.9 2 2 2h14l4 4-.01-18zM18 14H6v-2h12v2zm0-3H6V9h12v2zm0-3H6V6h12v2z"/></svg>',
    'folder': '<svg class="mui-icon" viewBox="0 0 24 24"><path d="M10 4H4c-1.1 0-1.99.9-1.99 2L2 18c0 1.1.9 2 2 2h16c1.1 0 2-.9 2-2V8c0-1.1-.9-2-2-2h-8l-2-2z"/></svg>',
    'play': '<svg class="mui-icon" viewBox="0 0 24 24"><path d="M8 5v14l11-7z"/></svg>',
    'pause': '<svg class="mui-icon" viewBox="0 0 24 24"><path d="M6 19h4V5H6v14zm8-14v14h4V5h-4z"/></svg>',
        'backup': '<svg class="mui-icon" viewBox="0 0 24 24"><path d="M19.35 10.04C18.67 6.59 15.64 4 12 4 9.11 4 6.6 5.64 5.35 8.04 2.34 8.36 0 10.91 0 14c0 3.31 2.69 6 6 6h13c2.76 0 5-2.24 5-5 0-2.64-2.05-4.78-4.65-4.96zM14 13v4h-4v-4H7l5-5 5 5h-3z"/></svg>',
    'calc': '<svg class="mui-icon" viewBox="0 0 24 24"><path d="M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm-6 2h5v2h-5V5zm0 3h5v2h-5V8zm-6-3h4v5H7V5zm0 7h4v2H7v-2zm0 3h4v2H7v-2zm6-3h5v2h-5v-2zm0 3h5v2h-5v-2z"/></svg>',
    'whatsapp': '<svg class="mui-icon" viewBox="0 0 24 24"><path d="M12.04 2C6.58 2 2.13 6.45 2.13 11.91c0 1.75.46 3.45 1.32 4.95L2.05 22l5.25-1.38c1.45.79 3.08 1.21 4.74 1.21 5.46 0 9.91-4.45 9.91-9.91 0-2.65-1.03-5.14-2.9-7.01A9.816 9.816 0 0 0 12.04 2zm.01 1.67c2.2 0 4.26.86 5.82 2.41a8.173 8.173 0 0 1 2.41 5.83c0 4.54-3.7 8.24-8.24 8.24-1.48 0-2.93-.4-4.2-1.15l-.3-.18-3.12.82.83-3.04-.2-.31a8.188 8.188 0 0 1-1.26-4.38c0-4.54 3.7-8.24 8.24-8.24zm4.52 11.66c-.25.7-.72 1.28-1.38 1.64-.66.36-1.44.42-2.22.18-.52-.16-1.2-.41-2.06-.91-1.28-.74-2.42-1.74-3.35-2.94-.93-1.2-1.48-2.52-1.63-3.92-.08-.77.12-1.52.55-2.11.43-.59 1.05-.93 1.74-.96.22-.01.44.04.64.14.2.1.36.26.47.46l.87 1.83c.1.21.12.45.06.67-.06.22-.2.4-.38.54l-.49.38c.34.69.83 1.3 1.43 1.8.6.5 1.3.87 2.06 1.09l.48-.48c.16-.16.36-.26.58-.29.22-.03.44.02.63.13l1.83.87c.2.1.36.26.46.46.1.2.13.43.08.65z"/></svg>',
    'fullscreen': '<svg class="mui-icon" viewBox="0 0 24 24"><path d="M7 14H5v5h5v-2H7v-3zm-2-4h2V7h3V5H5v5zm12 7h-3v2h5v-5h-2v3zM14 5v2h3v3h2V5h-5z"/></svg>',
    'history': '<svg class="mui-icon" viewBox="0 0 24 24"><path d="M13 3c-4.97 0-9 4.03-9 9H1l3.89 3.89.07.14L9 12H6c0-3.87 3.13-7 7-7s7 3.13 7 7-3.13 7-7 7c-1.93 0-3.68-.79-4.94-2.06l-1.42 1.42C8.27 19.99 10.51 21 13 21c4.97 0 9-4.03 9-9s-4.03-9-9-9zm-1 5v5l4.28 2.54.72-1.21-3.5-2.08V8H12z"/></svg>',
    'refresh': '<svg class="mui-icon" viewBox="0 0 24 24"><path d="M17.65 6.35C16.2 4.9 14.21 4 12 4c-4.42 0-7.99 3.58-7.99 8s3.57 8 7.99 8c3.73 0 6.84-2.55 7.73-6h-2.08c-.82 2.33-3.04 4-5.65 4-3.31 0-6-2.69-6-6s2.69-6 6-6c1.66 0 3.14.69 4.22 1.78L13 11h7V4l-2.35 2.35z"/></svg>',
    'download': '<svg class="mui-icon" viewBox="0 0 24 24"><path d="M19 9h-4V3H9v6H5l7 7 7-7zM5 18v2h14v-2H5z"/></svg>',
    'cloud': '<svg class="mui-icon" viewBox="0 0 24 24"><path d="M19.35 10.04C18.67 6.59 15.64 4 12 4 9.11 4 6.6 5.64 5.35 8.04 2.34 8.36 0 10.91 0 14c0 3.31 2.69 6 6 6h13c2.76 0 5-2.24 5-5 0-2.64-2.05-4.78-4.65-4.96zM19 18H6c-2.21 0-4-1.79-4-4 0-2.05 1.53-3.76 3.56-3.97l1.07-.11.5-.95C8.08 7.14 9.94 6 12 6c2.62 0 4.88 1.86 5.39 4.43l.3 1.5 1.53.11c1.56.1 2.78 1.41 2.78 2.96 0 1.65-1.35 3-3 3z"/></svg>',
    'cloud_sync': '<svg class="mui-icon" viewBox="0 0 24 24"><path d="M19.35 10.04C18.67 6.59 15.64 4 12 4 9.11 4 6.6 5.64 5.35 8.04 2.34 8.36 0 10.91 0 14c0 3.31 2.69 6 6 6h13c2.76 0 5-2.24 5-5 0-2.64-2.05-4.78-4.65-4.96zm-7.35 2.96V9l-4 4 4 4v-4h4v-2h-4z"/></svg>',
    'lock': '<svg class="mui-icon" viewBox="0 0 24 24"><path d="M18 8h-1V6c0-2.76-2.24-5-5-5S7 3.24 7 6v2H6c-1.1 0-2 .9-2 2v10c0 1.1.9 2 2 2h12c1.1 0 2-.9 2-2V10c0-1.1-.9-2-2-2zm-6 9c-1.1 0-2-.9-2-2s.9-2 2-2 2 .9 2 2-.9 2-2 2zm3.1-9H8.9V6c0-1.71 1.39-3.1 3.1-3.1 1.71 0 3.1 1.39 3.1 3.1v2z"/></svg>',
    'shield': '<svg class="mui-icon" viewBox="0 0 24 24"><path d="M12 1L3 5v6c0 5.55 3.84 10.74 9 12 5.16-1.26 9-6.45 9-12V5l-9-4zm0 10.99h7c-.53 4.12-3.28 7.79-7 8.94V12H5V6.3l7-3.11v8.8z"/></svg>',
    'settings': '<svg class="mui-icon" viewBox="0 0 24 24"><path d="M19.14 12.94c.04-.3.06-.61.06-.94 0-.32-.02-.64-.07-.94l2.03-1.58c.18-.14.23-.41.12-.61l-1.92-3.32c-.12-.22-.37-.29-.59-.22l-2.39.96c-.5-.38-1.03-.7-1.62-.94l-.36-2.54c-.04-.24-.24-.41-.48-.41h-3.84c-.24 0-.43.17-.47.41l-.36 2.54c-.59.24-1.13.57-1.62.94l-2.39-.96c-.22-.08-.47 0-.59.22L2.74 8.87c-.12.21-.08.47.12.61l2.03 1.58c-.05.3-.09.63-.09.94s.02.64.07.94l-2.03 1.58c-.18.14-.23.41-.12.61l1.92 3.32c.12.22.37.29.59.22l2.39-.96c.5.38 1.03.7 1.62.94l.36 2.54c.05.24.24.41.48.41h3.84c.24 0 .44-.17.47-.41l.36-2.54c.59-.24 1.13-.56 1.62-.94l2.39.96c.22.08.47 0 .59-.22l1.92-3.32c.12-.22.07-.47-.12-.61l-2.01-1.58zM12 15.6c-1.98 0-3.6-1.62-3.6-3.6s1.62-3.6 3.6-3.6 3.6 1.62 3.6 3.6-1.62 3.6-3.6 3.6z"/></svg>',
    'more_vert': '<svg class="mui-icon" viewBox="0 0 24 24"><path d="M12 8c1.1 0 2-.9 2-2s-.9-2-2-2-2 .9-2 2 .9 2 2 2zm0 2c-1.1 0-2 .9-2 2s.9 2 2 2 2-.9 2-2-.9-2-2-2zm0 6c-1.1 0-2 .9-2 2s.9 2 2 2 2-.9 2-2-.9-2-2-2z"/></svg>'
}
mui_icons_json = json.dumps(MUI_ICONS)


STUDIO_CSS = """
:root {
  --bg-main: #070A10;
  --bg-panel: #0D121D;
  --bg-card: #121826;
  --bg-card-hover: #182033;
  --bg-input: #0A0E17;
  --border: rgba(255, 255, 255, 0.08);
  --border-focus: #00E5FF;
  --teal: #00E5FF;
  --teal-glow: rgba(0, 229, 255, 0.25);
  --purple: #A78BFA;
  --purple-glow: rgba(167, 139, 250, 0.22);
  --emerald: #10B981;
  --rose: #F43F5E;
  --text-main: #F8FAFC;
  --text-muted: #94A3B8;
  --text-dim: #64748B;
  --radius: 14px;
}

* { box-sizing: border-box; margin: 0; padding: 0; }

/* App-Styled Custom Scrollbars */
::-webkit-scrollbar { width: 6px; height: 6px; }
::-webkit-scrollbar-track { background: #070A10; border-radius: 4px; }
::-webkit-scrollbar-thumb { background: rgba(255, 255, 255, 0.16); border-radius: 4px; }
::-webkit-scrollbar-thumb:hover { background: var(--teal); }

.mui-icon {
  width: 1.15em;
  height: 1.15em;
  fill: currentColor;
  display: inline-block;
  vertical-align: -0.2em;
  flex-shrink: 0;
  transition: transform 0.15s ease;
}
.btn-header .mui-icon,
.btn-add-action .mui-icon,
.header-tab-btn .mui-icon,
.btn-lib-insert .mui-icon {
  width: 1.25em;
  height: 1.25em;
  vertical-align: -0.22em;
}

* { scrollbar-width: thin; scrollbar-color: rgba(255, 255, 255, 0.16) #070A10; }

body {
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
  background: var(--bg-main);
  color: var(--text-main);
  height: 100vh;
  overflow: hidden;
  display: flex;
  flex-direction: column;
}

/* Master Top Navigation Header */
.master-header {
  height: 60px;
  background: var(--bg-panel);
  border-bottom: 1px solid var(--border);
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 0 24px;
  flex-shrink: 0;
  z-index: 100;
}
.brand-group {
  display: flex;
  align-items: center;
  gap: 16px;
}
.brand-logo {
  font-size: 1.25em;
  font-weight: 800;
  background: linear-gradient(135deg, var(--teal), var(--purple));
  -webkit-background-clip: text;
  -webkit-text-fill-color: transparent;
  display: flex;
  align-items: center;
  gap: 8px;
  cursor: pointer;
  user-select: none;
}
.nav-breadcrumb {
  display: flex;
  align-items: center;
  gap: 8px;
  font-size: 0.86em;
  color: var(--text-muted);
}
.btn-nav-crumb {
  height: 38px;
  box-sizing: border-box;
  background: rgba(255, 255, 255, 0.04);
  border: 1px solid var(--border);
  color: var(--text-muted);
  font-size: 0.85em;
  font-weight: 700;
  padding: 0 14px;
  border-radius: 9px;
  cursor: pointer;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  gap: 8px;
  transition: 0.15s all;
  white-space: nowrap;
}
.btn-nav-crumb:hover {
  background: rgba(255, 255, 255, 0.08);
  color: var(--text-main);
  border-color: rgba(255, 255, 255, 0.2);
}
.btn-nav-crumb:hover { background: rgba(255, 255, 255, 0.1); color: var(--text-main); border-color: rgba(255, 255, 255, 0.25); }

.header-nav-tabs {
  display: flex;
  align-items: center;
  gap: 6px;
  margin-left: 16px;
  background: rgba(255, 255, 255, 0.03);
  padding: 4px;
  border-radius: 10px;
  border: 1px solid var(--border);
}
.header-tab-btn {
  background: transparent;
  border: 1px solid transparent;
  color: var(--text-muted);
  font-size: 0.82em;
  font-weight: 700;
  padding: 6px 14px;
  border-radius: 8px;
  cursor: pointer;
  display: flex;
  align-items: center;
  gap: 6px;
  transition: 0.15s all;
  user-select: none;
}
.header-tab-btn:hover:not(.active) {
  color: var(--text-main);
  background: rgba(255, 255, 255, 0.06);
}
.header-tab-btn.active {
  background: linear-gradient(135deg, var(--teal), #00A3FF) !important;
  color: #05080E !important;
  border: none !important;
  font-weight: 800;
  box-shadow: 0 4px 14px var(--teal-glow);
}

.header-right-actions {
  display: flex;
  align-items: center;
  gap: 10px;
}


/* COMPACT HEADER SWITCHER & STUDIO MENU */
.header-switcher-wrap {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  position: relative;
}
.header-switcher-sep {
  color: var(--text-dim);
  font-size: 0.9em;
  font-weight: 700;
  user-select: none;
  margin: 0 2px;
}
.switcher-btn {
  background: rgba(255, 255, 255, 0.05);
  border: 1px solid var(--border);
  color: var(--text-main);
  padding: 6px 12px;
  border-radius: 8px;
  font-size: 0.84em;
  font-weight: 700;
  cursor: pointer;
  display: inline-flex;
  align-items: center;
  gap: 6px;
  transition: 0.15s all;
  white-space: nowrap;
}
.switcher-btn:hover, .switcher-btn.active {
  background: var(--bg-card-hover);
  border-color: var(--teal);
  box-shadow: 0 0 10px var(--teal-glow);
  color: var(--teal);
}
.switcher-menu {
  position: absolute;
  top: calc(100% + 6px);
  left: 0;
  min-width: 200px;
  background: #0D121D;
  border: 1px solid rgba(0, 229, 255, 0.35);
  border-radius: 12px;
  box-shadow: 0 14px 40px rgba(0, 0, 0, 0.8), 0 0 16px var(--teal-glow);
  padding: 6px;
  z-index: 9999;
  display: none;
  max-height: 280px;
  overflow-y: auto;
  backdrop-filter: blur(16px);
}
.switcher-menu.open {
  display: block;
}
.btn-header.icon-only {
  padding: 8px 10px;
  min-width: 36px;
  justify-content: center;
}

/* CUSTOM APP-STYLED SELECT DROPDOWNS */
.custom-select-wrapper {
  position: relative;
  display: inline-block;
  vertical-align: middle;
  min-width: 150px;
}
.custom-select-trigger {
  width: 100%;
  height: 38px !important;
  box-sizing: border-box !important;
  background: var(--bg-card) !important;
  border: 1px solid var(--border) !important;
  color: var(--text-main) !important;
  padding: 0 12px !important;
  border-radius: 9px !important;
  font-size: 0.85em !important;
  font-weight: 700 !important;
  display: flex !important;
  align-items: center !important;
  justify-content: space-between !important;
  gap: 8px !important;
  cursor: pointer !important;
  transition: 0.18s all ease !important;
  user-select: none !important;
  white-space: nowrap !important;
}
.custom-select-trigger:hover,
.custom-select-trigger.active {
  border-color: var(--teal) !important;
  box-shadow: 0 0 12px var(--teal-glow);
  background: var(--bg-card-hover);
}
.custom-select-arrow {
  color: var(--text-muted);
  display: flex;
  align-items: center;
  transition: transform 0.2s ease;
}
.custom-select-trigger.active .custom-select-arrow {
  transform: rotate(180deg);
  color: var(--teal);
}
.custom-select-menu {
  position: absolute;
  top: calc(100% + 6px);
  left: 0;
  min-width: 100%;
  width: max-content;
  background: #0D121D;
  border: 1px solid rgba(0, 229, 255, 0.35);
  border-radius: 12px;
  box-shadow: 0 14px 40px rgba(0, 0, 0, 0.8), 0 0 16px var(--teal-glow);
  padding: 8px;
  z-index: 9999;
  display: none;
  opacity: 0;
  transform: translateY(-6px);
  transition: opacity 0.15s ease, transform 0.15s ease;
  max-height: 240px;
  overflow-y: auto;
  box-sizing: border-box;
  backdrop-filter: blur(16px);
}
.custom-select-menu.open {
  display: flex;
  flex-direction: column;
  gap: 8px;
  opacity: 1;
  transform: translateY(0);
}
.custom-select-item {
  padding: 10px 14px;
  border-radius: 8px;
  color: var(--text-main);
  white-space: nowrap;
  font-size: 0.86em;
  font-weight: 700;
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 14px;
  cursor: pointer;
  transition: background-color 0.15s ease, color 0.15s ease;
  user-select: none;
}
.custom-select-item:hover {
  background: rgba(0, 229, 255, 0.12);
  color: var(--teal);
}
.custom-select-item.selected {
  background: rgba(0, 229, 255, 0.18);
  color: var(--teal);
  font-weight: 800;
}
.custom-select-item.selected:hover {
  background: rgba(0, 229, 255, 0.24);
}
.custom-select-item .mui-icon {
  width: 16px;
  height: 16px;
  color: var(--teal);
}

.btn-header {
  height: 38px;
  box-sizing: border-box;
  background: rgba(255, 255, 255, 0.04);
  border: 1px solid var(--border);
  color: var(--text-main);
  font-size: 0.85em;
  font-weight: 700;
  padding: 0 14px;
  border-radius: 9px;
  cursor: pointer;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  gap: 8px;
  transition: 0.2s all;
  white-space: nowrap;
  flex-shrink: 0;
}
.btn-header:hover { background: rgba(255, 255, 255, 0.12); border-color: rgba(255, 255, 255, 0.25); color: #FFF; }
.btn-header.btn-header-primary {
  background: linear-gradient(135deg, var(--teal), #00A3FF) !important;
  color: #05080E !important;
  border: none !important;
  font-weight: 800;
  box-shadow: 0 4px 16px var(--teal-glow);
}
.btn-header.btn-header-primary:hover {
  background: linear-gradient(135deg, #4DF0FF, #38B6FF) !important;
  color: #000 !important;
  box-shadow: 0 6px 24px rgba(0, 229, 255, 0.5) !important;
  transform: translateY(-1px);
}

/* Viewport */
.view-viewport {
  flex: 1;
  overflow: hidden;
  position: relative;
  display: flex;
  flex-direction: column;
}
.view-screen {
  position: absolute;
  top: 0; left: 0; right: 0; bottom: 0;
  display: none;
  flex-direction: column;
  overflow-y: auto;
  background: var(--bg-main);
  opacity: 0;
  pointer-events: none;
  transition: opacity 0.18s ease;
}
.view-screen.active {
  display: flex;
  opacity: 1;
  pointer-events: auto;
  z-index: 10;
}
.view-screen.active {
  opacity: 1;
  pointer-events: auto;
}

/* VIEW 1: ATHLETES ROSTER (HUB) */
.hub-content {
  max-width: 1080px;
  width: 100%;
  margin: 0 auto;
  padding: 28px 24px 80px 24px;
}
.hub-header-bar {
  display: flex;
  justify-content: space-between;
  align-items: center;
  background: var(--bg-panel);
  border: 1px solid var(--border);
  border-radius: var(--radius);
  padding: 20px 24px;
  margin-bottom: 24px;
}
.hub-header-info {
  display: flex;
  flex-direction: column;
}
.hub-header-actions {
  display: flex;
  gap: 10px;
  align-items: center;
}
.hub-search-input {
  width: 220px;
  padding: 8px 12px;
}
.athlete-profile-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 20px;
}
.athlete-profile-info {
  display: flex;
  align-items: center;
  gap: 14px;
}
.athlete-profile-actions {
  display: flex;
  gap: 8px;
  align-items: center;
}
.soc-header-row {
  display: flex;
  align-items: center;
  gap: 10px;
}
.soc-secondary-actions {
  display: flex;
  align-items: center;
  gap: 6px;
}
.builder-tb-top, .builder-tb-mid, .builder-tb-bottom {
  display: contents;
}
.hub-header-info {
  display: flex;
  flex-direction: column;
}
.hub-header-actions {
  display: flex;
  gap: 10px;
  align-items: center;
}
.athlete-profile-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 20px;
}
.athlete-profile-info {
  display: flex;
  align-items: center;
  gap: 14px;
}
.athlete-profile-actions {
  display: flex;
  gap: 8px;
  align-items: center;
}
.soc-header-row {
  display: flex;
  align-items: center;
  gap: 10px;
}
.soc-secondary-actions {
  display: flex;
  align-items: center;
  gap: 6px;
}
.builder-tb-top, .builder-tb-mid, .builder-tb-bottom {
  display: contents;
}

.athlete-badge-count {
  font-size: 0.68em;
  background: var(--teal-glow);
  color: var(--teal);
  border: 1px solid rgba(0, 229, 255, 0.3);
  padding: 3px 10px;
  border-radius: 20px;
  font-weight: 800;
}
.athletes-grid {
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(320px, 1fr));
  gap: 20px;
}
.athlete-roster-card {
  background: var(--bg-panel);
  border: 1px solid var(--border);
  border-radius: 16px;
  padding: 22px;
  display: flex;
  flex-direction: column;
  justify-content: space-between;
  transition: 0.2s all;
  cursor: pointer;
}
.athlete-roster-card:hover {
  border-color: rgba(0, 229, 255, 0.35);
  transform: translateY(-2px);
  box-shadow: 0 12px 35px rgba(0, 0, 0, 0.6);
}
.athlete-card-top {
  display: flex;
  align-items: center;
  gap: 14px;
  margin-bottom: 14px;
}
.athlete-avatar-circle {
  width: 48px;
  height: 48px;
  border-radius: 14px;
  background: linear-gradient(135deg, rgba(0, 229, 255, 0.15), rgba(167, 139, 250, 0.15));
  border: 1px solid var(--teal);
  color: var(--teal);
  font-weight: 800;
  font-size: 1.1em;
  display: flex;
  align-items: center;
  justify-content: center;
  flex-shrink: 0;
}
.athlete-card-name {
  font-size: 1.15em;
  font-weight: 800;
  color: var(--text-main);
}
.athlete-card-sessions-count {
  font-size: 0.78em;
  color: var(--purple);
  font-weight: 700;
  margin-top: 2px;
}
.athlete-card-comments {
  background: var(--bg-card);
  border-radius: 10px;
  padding: 10px 12px;
  font-size: 0.82em;
  color: var(--text-muted);
  line-height: 1.4;
  margin-bottom: 16px;
  border: 1px solid rgba(255, 255, 255, 0.04);
  max-height: 60px;
  overflow: hidden;
  text-overflow: ellipsis;
}
.athlete-card-footer {
  display: flex;
  gap: 8px;
  justify-content: space-between;
  align-items: center;
}

/* VIEW 2: ATHLETE DETAIL / SESSIONS OVERVIEW */
.athlete-detail-content {
  max-width: 1040px;
  width: 100%;
  margin: 0 auto;
  padding: 24px 20px 80px 20px;
}
.athlete-comments-card {
  background: var(--bg-panel);
  border: 1px solid var(--border);
  border-radius: var(--radius);
  padding: 18px 22px;
  margin-bottom: 24px;
}
.comments-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 8px;
}
.comments-title {
  font-size: 0.86em;
  font-weight: 800;
  text-transform: uppercase;
  color: var(--purple);
  display: flex;
  align-items: center;
  gap: 6px;
}
.comments-textarea {
  width: 100%;
  background: var(--bg-input);
  border: 1px solid var(--border);
  border-radius: 10px;
  color: var(--text-main);
  font-family: inherit;
  font-size: 0.9em;
  line-height: 1.5;
  padding: 12px 14px;
  outline: none;
  min-height: 75px;
  resize: vertical;
}
.comments-textarea:focus { border-color: var(--teal); }

.sessions-manager-bar {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 16px;
}
.sessions-manager-title {
  font-size: 1.2em;
  font-weight: 800;
  color: var(--text-main);
  display: flex;
  align-items: center;
  gap: 10px;
}

.sessions-overview-grid {
  display: flex;
  flex-direction: column;
  gap: 12px;
}
.session-overview-card {
  background: var(--bg-card);
  border: 1px solid var(--border);
  border-radius: 14px;
  padding: 16px 20px;
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 16px;
  transition: 0.2s all;
}
.session-overview-card:hover {
  background: var(--bg-card-hover);
  border-color: rgba(255, 255, 255, 0.16);
  transform: translateY(-1px);
}
.soc-left {
  display: flex;
  align-items: center;
  gap: 14px;
  flex: 1;
  min-width: 0;
}
.soc-num {
  font-size: 0.74em;
  font-weight: 800;
  background: rgba(0, 229, 255, 0.12);
  color: var(--teal);
  padding: 4px 9px;
  border-radius: 8px;
}
.soc-title-wrap {
  flex: 1;
  min-width: 0;
}
.soc-title {
  font-size: 1.05em;
  font-weight: 800;
  color: var(--text-main);
  margin-bottom: 4px;
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
}
.soc-badges {
  display: flex;
  align-items: center;
  gap: 8px;
  font-size: 0.74em;
  color: var(--text-muted);
}
.soc-badge {
  background: rgba(255, 255, 255, 0.05);
  padding: 2px 7px;
  border-radius: 6px;
}
.soc-preview-exercises {
  font-size: 0.78em;
  color: var(--text-dim);
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
  margin-top: 4px;
}
.soc-actions {
  display: flex;
  align-items: center;
  gap: 6px;
  flex-shrink: 0;
}

/* VIEW 3: WORKOUT BUILDER / EXERCISE EDITOR */
.builder-content {
  max-width: 980px;
  width: 100%;
  margin: 0 auto;
  padding: 24px 20px 80px 20px;
}
.exercise-row {
  background: var(--bg-card);
  border: 1px solid var(--border);
  border-radius: 12px;
  padding: 12px 16px;
  margin-bottom: 10px;
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  transition: 0.2s all;
  cursor: pointer;
}
.exercise-row:hover {
  background: var(--bg-card-hover);
  border-color: rgba(255, 255, 255, 0.18);
  transform: translateY(-1px);
}
.exercise-row.dragging { opacity: 0.35; }
.row-left {
  display: flex;
  align-items: center;
  gap: 12px;
  flex: 1;
  min-width: 0;
}
.drag-handle {
  cursor: grab;
  color: var(--text-dim);
  font-size: 1.1em;
  user-select: none;
  padding: 2px 4px;
}
.drag-handle:active { cursor: grabbing; }
.row-num {
  font-size: 0.72em;
  font-weight: 800;
  background: rgba(0, 229, 255, 0.12);
  color: var(--teal);
  padding: 3px 8px;
  border-radius: 6px;
}
.row-cat-badge {
  font-size: 0.7em;
  font-weight: 800;
  text-transform: uppercase;
  color: var(--purple);
  background: var(--purple-glow);
  padding: 3px 8px;
  border-radius: 6px;
}
.row-title {
  font-weight: 700;
  font-size: 0.98em;
  color: var(--text-main);
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
}
.row-pill-info {
  display: flex;
  align-items: center;
  gap: 8px;
  font-size: 0.78em;
  color: var(--text-muted);
}
.row-badge {
  background: rgba(255, 255, 255, 0.05);
  padding: 3px 8px;
  border-radius: 6px;
  color: var(--text-muted);
}
.row-actions {
  display: flex;
  align-items: center;
  gap: 4px;
  flex-shrink: 0;
}
.btn-row-action {
  background: transparent;
  border: 1px solid transparent;
  color: var(--text-muted);
  width: 32px;
  height: 32px;
  border-radius: 7px;
  display: flex;
  align-items: center;
  justify-content: center;
  cursor: pointer;
  font-size: 0.88em;
  transition: 0.15s all;
}
.btn-row-action:hover { background: rgba(255, 255, 255, 0.08); color: var(--text-main); }
.btn-row-action.edit-btn {
  background: rgba(0, 229, 255, 0.08);
  color: var(--teal);
  border: 1px solid rgba(0, 229, 255, 0.2);
  width: auto;
  padding: 0 10px;
  font-weight: 700;
  font-size: 0.78em;
  gap: 4px;
}
.btn-row-action.edit-btn:hover { background: var(--teal); color: #000; }
.btn-row-action.danger:hover { background: rgba(244, 63, 94, 0.15); color: var(--rose); }

.add-bar { display: flex; gap: 12px; margin-top: 18px; }
.btn-add-action {
  flex: 1;
  padding: 14px;
  border-radius: 12px;
  font-size: 0.88em;
  font-weight: 800;
  cursor: pointer;
  border: 1px dashed var(--border);
  background: rgba(255, 255, 255, 0.02);
  color: var(--text-main);
  display: flex;
  align-items: center;
  justify-content: center;
  gap: 8px;
  transition: 0.2s all;
}
.btn-add-action:hover { background: rgba(0, 229, 255, 0.06); border-color: var(--teal); color: var(--teal); }
.btn-add-action.lib-btn {
  border-color: rgba(167, 139, 250, 0.35);
  background: rgba(167, 139, 250, 0.05);
  color: var(--purple);
}
.btn-add-action.lib-btn:hover { background: rgba(167, 139, 250, 0.12); border-color: var(--purple); }

/* APP-STYLED IN-APP TOAST NOTIFICATION */
.toast-notification {
  position: fixed;
  bottom: 28px;
  left: 50%;
  transform: translateX(-50%) translateY(90px);
  background: #0E1422;
  border: 1px solid var(--teal);
  box-shadow: 0 10px 35px rgba(0, 229, 255, 0.28);
  color: #FFF;
  padding: 12px 24px;
  border-radius: 30px;
  font-size: 0.88em;
  font-weight: 700;
  z-index: 9999;
  opacity: 0;
  transition: 0.28s all cubic-bezier(0.16, 1, 0.3, 1);
  pointer-events: none;
  display: flex;
  align-items: center;
  gap: 8px;
}
.toast-notification.show {
  transform: translateX(-50%) translateY(0);
  opacity: 1;
}

/* APP-STYLED GLASSMORPHIC MODALS */
.modal-overlay {
  position: fixed; top: 0; left: 0; right: 0; bottom: 0;
  background: rgba(4, 7, 12, 0.88); backdrop-filter: blur(14px);
  display: flex; justify-content: center; align-items: center;
  padding: 20px; z-index: 5000; opacity: 0; pointer-events: none;
  transition: 0.22s all cubic-bezier(0.16, 1, 0.3, 1);
}
.modal-overlay.show { opacity: 1; pointer-events: auto; }
.modal-window {
  background: #0E1422; border: 1px solid var(--border);
  border-top: 1px solid rgba(255, 255, 255, 0.18);
  border-radius: 24px; padding: 26px; max-width: 640px; width: 95%; max-height: 90vh;
  display: flex; flex-direction: column; box-shadow: 0 30px 90px rgba(0,0,0,0.9);
  transform: scale(0.95); transition: 0.22s all cubic-bezier(0.16, 1, 0.3, 1);
  overflow-y: auto;
}
.modal-overlay.show .modal-window { transform: scale(1); }
.modal-header-bar {
  display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px; flex-shrink: 0;
}
.modal-title { font-size: 1.25em; font-weight: 800; color: var(--text-main); display: flex; align-items: center; gap: 8px; }
.modal-close-btn {
  background: rgba(255, 255, 255, 0.06); border: 1px solid var(--border);
  color: var(--text-muted); width: 32px; height: 32px; border-radius: 50%;
  cursor: pointer; font-size: 0.9em; display: flex; align-items: center; justify-content: center;
  transition: 0.15s all;
}
.modal-close-btn:hover { background: rgba(255, 255, 255, 0.12); color: var(--text-main); }

.form-row-2 { display: grid; grid-template-columns: 1fr 1fr; gap: 14px; margin-bottom: 14px; }
.form-row-3 { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 12px; margin-bottom: 14px; }
.form-group { display: flex; flex-direction: column; gap: 6px; margin-bottom: 14px; }
.form-label { font-size: 0.72em; font-weight: 800; text-transform: uppercase; color: var(--text-muted); }
.form-input {
  background: var(--bg-input); border: 1px solid var(--border);
  color: var(--text-main); font-size: 0.9em; padding: 10px 14px; border-radius: 8px; outline: none; width: 100%;
}
.form-input:focus { border-color: var(--teal); }
.form-textarea {
  background: var(--bg-input); border: 1px solid var(--border);
  color: var(--text-main); font-size: 0.88em; padding: 10px 14px; border-radius: 8px; outline: none;
  resize: vertical; min-height: 80px; width: 100%; font-family: inherit; line-height: 1.4;
}
.form-textarea:focus { border-color: var(--teal); }
.modal-footer-bar {
  display: flex; justify-content: flex-end; gap: 10px; margin-top: 20px; flex-shrink: 0;
}

/* Phone Simulator Modal (Auto-fit proportional scale) */
#preview-modal {
  align-items: center;
  justify-content: center;
  padding: 14px;
  overflow: hidden;
}
.preview-modal-window {
  background: transparent !important;
  border: none !important;
  padding: 0 !important;
  max-width: 440px;
  width: 100%;
  display: flex;
  flex-direction: column;
  align-items: center;
  box-shadow: none !important;
  overflow: visible !important;
  max-height: 98vh;
  width: auto;
}
.preview-modal-controls {
  display: flex;
  align-items: center;
  justify-content: space-between;
  width: 100%;
  max-width: 440px;
  margin-bottom: 14px;
  position: relative;
  z-index: 20;
}
.preview-stage {
  display: flex;
  justify-content: center;
  align-items: flex-start;
  position: relative;
  overflow: visible;
  margin: 0 auto;
}
.phone-chassis {
  width: 390px;
  height: 844px;
  background: #000000;
  border: 12px solid #1E2536;
  border-radius: 54px;
  box-shadow: 0 0 0 2px #334155, 0 0 0 5px #070B12, 0 25px 80px rgba(0, 0, 0, 0.95), 0 0 40px rgba(0, 229, 255, 0.08);
  position: absolute;
  top: 0;
  left: 50%;
  overflow: hidden;
  display: flex;
  flex-direction: column;
  box-sizing: content-box;
  transform-origin: top center;
  transform: translateX(-50%);
  transition: width 0.25s ease, height 0.25s ease, transform 0.25s ease;
  flex-shrink: 0;
}
.phone-chassis.wide {
  width: 430px;
  height: 932px;
  border-radius: 58px;
}
.phone-notch {
  position: absolute;
  top: 12px;
  left: 50%;
  transform: translateX(-50%);
  width: 96px;
  height: 24px;
  background: #000;
  border-radius: 20px;
  z-index: 1000;
  pointer-events: none;
  box-shadow: 0 0 0 1px rgba(255, 255, 255, 0.08);
  display: flex;
  align-items: center;
  justify-content: flex-end;
  padding-right: 8px;
  box-sizing: border-box;
}
.phone-notch::after {
  content: "";
  width: 9px;
  height: 9px;
  border-radius: 50%;
  background: #0E1626;
  box-shadow: inset 0 0 2px rgba(0, 229, 255, 0.35);
}
.client-iframe {
  width: 100%;
  height: 100%;
  border: none;
  background: #070A10;
  border-radius: 42px;
  overflow: hidden;
  display: block;
}
.phone-chassis.wide .client-iframe {
  border-radius: 46px;
}

/* Library Modal */
.lib-modal-window {
  background: #0E1422;
  border: 1px solid var(--border);
  border-radius: 24px;
  padding: 24px;
  max-width: 860px;
  width: 95%;
  max-height: 88vh;
  display: flex;
  flex-direction: column;
  box-shadow: 0 30px 90px rgba(0,0,0,0.9);
}
.lib-search-input {
  width: 100%;
  background: var(--bg-input);
  border: 1px solid var(--border);
  color: var(--text-main);
  padding: 11px 16px;
  border-radius: 10px;
  font-size: 0.92em;
  margin-bottom: 12px;
  outline: none;
}
.lib-search-input:focus {
  border-color: var(--teal);
  box-shadow: 0 0 12px var(--teal-glow);
}
.lib-pills-row {
  display: flex;
  gap: 8px;
  overflow-x: auto;
  padding: 4px 2px 12px 2px;
  margin-bottom: 6px;
  flex-shrink: 0;
  scrollbar-width: none;
}
.lib-pills-row::-webkit-scrollbar {
  display: none;
}
.lib-filter-pill {
  background: rgba(255, 255, 255, 0.05);
  border: 1px solid var(--border);
  color: var(--text-muted);
  font-size: 0.8em;
  font-weight: 700;
  padding: 8px 18px;
  border-radius: 20px;
  cursor: pointer;
  white-space: nowrap;
  display: inline-flex;
  align-items: center;
  height: 36px;
  box-sizing: border-box;
  transition: 0.15s all;
  user-select: none;
}
.lib-filter-pill:hover:not(.active) {
  background: rgba(255, 255, 255, 0.1);
  color: var(--text-main);
  border-color: rgba(255, 255, 255, 0.22);
}
.lib-filter-pill.active {
  background: linear-gradient(135deg, var(--teal), #00A3FF) !important;
  color: #070A10 !important;
  font-weight: 800 !important;
  border-color: transparent !important;
  box-shadow: 0 4px 16px var(--teal-glow) !important;
}
.lib-cards-grid {
  flex: 1;
  overflow-y: auto;
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(250px, 1fr));
  gap: 14px;
  padding: 8px 8px 24px 8px;
}
.lib-item-card {
  background: var(--bg-card);
  border: 1.5px solid var(--border);
  border-radius: 14px;
  padding: 16px;
  display: flex;
  flex-direction: column;
  justify-content: space-between;
  gap: 10px;
  transition: 0.18s border-color, 0.18s box-shadow, 0.18s background;
  box-sizing: border-box;
}
.lib-item-card:hover {
  border-color: var(--teal) !important;
  box-shadow: 0 0 16px var(--teal-glow);
  background: var(--bg-card-hover);
}
.btn-lib-insert {
  background: rgba(0, 229, 255, 0.1);
  border: 1px solid rgba(0, 229, 255, 0.3);
  color: var(--teal);
  border-radius: 9px;
  padding: 9px 12px;
  font-size: 0.82em;
  font-weight: 800;
  cursor: pointer;
  width: 100%;
  transition: 0.15s all;
  display: flex;
  align-items: center;
  justify-content: center;
  gap: 6px;
}
.btn-lib-insert:hover {
  background: var(--teal);
  color: #000;
  box-shadow: 0 0 12px var(--teal-glow);
}

/* Export Sessions Modal Checklist */
.export-sessions-container {
  scrollbar-width: thin;
  scrollbar-color: rgba(255, 255, 255, 0.15) transparent;
}
.export-sessions-container::-webkit-scrollbar {
  width: 4px;
}
.export-sessions-container::-webkit-scrollbar-thumb {
  background: rgba(255, 255, 255, 0.15);
  border-radius: 4px;
}
.export-session-item {
  border: 1px solid rgba(255, 255, 255, 0.08) !important;
  background: rgba(255, 255, 255, 0.03) !important;
  transition: all 0.18s ease !important;
}
.export-session-item:hover {
  background: rgba(255, 255, 255, 0.06) !important;
  border-color: rgba(0, 229, 255, 0.25) !important;
}
.export-session-item.selected {
  background: rgba(0, 229, 255, 0.06) !important;
  border-color: rgba(0, 229, 255, 0.35) !important;
}

/* Builder Toolbar Unified Layout */
.builder-toolbar {
  display: flex;
  align-items: center;
  gap: 10px;
  background: var(--bg-panel);
  border: 1px solid var(--border);
  border-radius: var(--radius);
  padding: 10px 14px;
  margin-bottom: 20px;
  position: relative;
  z-index: 50;
  overflow: visible;
}
.toolbar-divider {
  width: 1px;
  height: 22px;
  background: rgba(255, 255, 255, 0.1);
  margin: 0 2px;
  flex-shrink: 0;
}
.builder-session-input {
  height: 38px !important;
  box-sizing: border-box !important;
  background: var(--bg-input) !important;
  border: 1px solid var(--border) !important;
  color: var(--text-main) !important;
  font-size: 0.95em !important;
  font-weight: 700 !important;
  padding: 0 12px !important;
  border-radius: 9px !important;
  outline: none !important;
  width: 180px !important;
  min-width: 130px !important;
  flex-shrink: 1 !important;
  transition: border-color 0.18s ease, box-shadow 0.18s ease !important;
}
.builder-session-input:focus {
  border-color: var(--teal) !important;
  box-shadow: 0 0 12px var(--teal-glow) !important;
}

/* ==========================================================================
   ATHLETE TRAINING CALENDAR & LOGS
   ========================================================================== */
.calendar-controls-bar {
  display: flex;
  align-items: center;
  justify-content: space-between;
  background: var(--bg-card);
  border: 1px solid var(--border);
  border-radius: var(--radius);
  padding: 12px 18px;
  margin-bottom: 18px;
  flex-wrap: wrap;
  gap: 12px;
}
.cal-nav-group {
  display: flex;
  align-items: center;
  gap: 8px;
}
.cal-month-title {
  font-size: 1.15em;
  font-weight: 800;
  color: var(--text-main);
  min-width: 170px;
  text-align: center;
  letter-spacing: -0.01em;
}
.cal-filter-group {
  display: flex;
  align-items: center;
  gap: 12px;
  flex-wrap: wrap;
}
.cal-stats-pill {
  font-size: 0.82em;
  font-weight: 700;
  color: var(--teal);
  background: rgba(0, 229, 255, 0.08);
  border: 1px solid rgba(0, 229, 255, 0.25);
  padding: 5px 14px;
  border-radius: 20px;
  white-space: nowrap;
}
.calendar-container {
  background: var(--bg-card);
  border: 1px solid var(--border);
  border-radius: 16px;
  overflow: hidden;
  display: flex;
  flex-direction: column;
}
.calendar-weekdays-grid {
  display: grid;
  grid-template-columns: repeat(7, 1fr);
  background: rgba(255, 255, 255, 0.02);
  border-bottom: 1px solid var(--border);
  padding: 12px 0;
  text-align: center;
  font-size: 0.78em;
  font-weight: 800;
  text-transform: uppercase;
  letter-spacing: 0.06em;
  color: var(--text-dim);
}
.calendar-days-grid {
  display: grid;
  grid-template-columns: repeat(7, 1fr);
  background: var(--border);
  gap: 1px;
}
.calendar-day-cell {
  background: var(--bg-card);
  min-height: 120px;
  padding: 8px 10px;
  display: flex;
  flex-direction: column;
  transition: background 0.18s ease;
  cursor: pointer;
  position: relative;
}
.calendar-day-cell:hover {
  background: var(--bg-card-hover);
}
.calendar-day-cell.other-month {
  background: rgba(13, 18, 29, 0.45);
  opacity: 0.45;
}
.calendar-day-cell.today {
  box-shadow: inset 0 0 0 1.5px var(--teal);
}
.cal-day-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 6px;
}
.cal-day-number {
  font-size: 0.85em;
  font-weight: 800;
  color: var(--text-muted);
}
.calendar-day-cell.today .cal-day-number {
  color: var(--teal);
}
.cal-add-log-btn {
  opacity: 0;
  transition: opacity 0.15s ease, background 0.15s ease;
  font-size: 0.72em;
  font-weight: 700;
  padding: 2px 7px;
  border-radius: 5px;
  background: rgba(255, 255, 255, 0.08);
  color: var(--text-main);
  border: none;
  cursor: pointer;
}
.cal-add-log-btn:hover {
  background: var(--teal);
  color: #070A10;
}
.calendar-day-cell:hover .cal-add-log-btn {
  opacity: 1;
}
.cal-badges-container {
  display: flex;
  flex-direction: column;
  gap: 4px;
  overflow-y: auto;
  max-height: 95px;
}
.cal-session-badge {
  font-size: 0.73em;
  padding: 4px 7px;
  border-radius: 6px;
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 4px;
  line-height: 1.25;
  font-weight: 600;
  transition: transform 0.15s ease, filter 0.15s ease;
  cursor: pointer;
}
.cal-session-badge:hover {
  transform: translateY(-1px);
  filter: brightness(1.15);
}
.cal-session-badge.completed {
  background: rgba(16, 185, 129, 0.15);
  color: #34d399;
  border: 1px solid rgba(16, 185, 129, 0.32);
}
.cal-session-badge.scheduled {
  background: rgba(0, 229, 255, 0.12);
  color: var(--teal);
  border: 1px solid rgba(0, 229, 255, 0.3);
}
.cal-session-badge.missed {
  background: rgba(244, 63, 94, 0.14);
  color: var(--rose);
  border: 1px solid rgba(244, 63, 94, 0.3);
}
.cal-badge-ath-avatar {
  width: 16px;
  height: 16px;
  border-radius: 50%;
  font-size: 9px;
  font-weight: 800;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  background: rgba(255, 255, 255, 0.18);
  color: #fff;
  flex-shrink: 0;
}

/* Athlete Profile History Section */
.athlete-history-card {
  background: var(--bg-card);
  border: 1px solid var(--border);
  border-radius: var(--radius);
  padding: 20px;
  margin-top: 24px;
}
.history-card-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 16px;
}
.history-timeline-list {
  display: flex;
  flex-direction: column;
  gap: 10px;
}
.history-item-row {
  display: flex;
  align-items: flex-start;
  justify-content: space-between;
  gap: 14px;
  padding: 12px 14px;
  background: rgba(255, 255, 255, 0.02);
  border: 1px solid var(--border);
  border-radius: 10px;
  transition: background 0.18s ease;
}
.history-item-row:hover {
  background: rgba(255, 255, 255, 0.04);
}
.history-date-badge {
  font-size: 0.78em;
  font-weight: 700;
  color: var(--teal);
  background: rgba(0, 229, 255, 0.1);
  border: 1px solid rgba(0, 229, 255, 0.22);
  padding: 4px 8px;
  border-radius: 6px;
  white-space: nowrap;
}
.history-item-title {
  font-size: 0.92em;
  font-weight: 800;
  color: var(--text-main);
}
.history-item-summary {
  font-size: 0.8em;
  color: var(--text-muted);
  margin-top: 3px;
}
.history-item-notes {
  font-size: 0.78em;
  color: var(--text-dim);
  margin-top: 4px;
  font-style: italic;
}

/* Modal Status Selector Pills */
.status-pill-group {
  display: flex;
  gap: 8px;
}
.status-pill-btn {
  flex: 1;
  padding: 7px 10px;
  border-radius: 8px;
  border: 1px solid var(--border);
  background: var(--bg-input);
  color: var(--text-muted);
  font-size: 0.82em;
  font-weight: 700;
  cursor: pointer;
  transition: all 0.15s ease;
  text-align: center;
}
.status-pill-btn.active.completed {
  background: rgba(16, 185, 129, 0.2);
  border-color: #10B981;
  color: #34d399;
}
.status-pill-btn.active.scheduled {
  background: rgba(0, 229, 255, 0.18);
  border-color: var(--teal);
  color: var(--teal);
}
.status-pill-btn.active.missed {
  background: rgba(244, 63, 94, 0.18);
  border-color: var(--rose);
  color: var(--rose);
}

/* Calendar Day Workout Inspector & Exercise Breakdown */
.cal-day-modal-window {
  max-width: 680px !important;
  max-height: 90vh !important;
  display: flex !important;
  flex-direction: column !important;
}
.cal-workout-card {
  background: var(--bg-card);
  border: 1px solid var(--border);
  border-radius: 14px;
  padding: 16px;
  margin-bottom: 14px;
  transition: border-color 0.18s ease;
}
.cal-workout-card:hover {
  border-color: rgba(0, 229, 255, 0.3);
}
.cal-workout-card-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 12px;
  padding-bottom: 10px;
  border-bottom: 1px solid var(--border);
}
.cal-exercise-item {
  display: flex;
  justify-content: space-between;
  align-items: center;
  background: rgba(255, 255, 255, 0.02);
  border: 1px solid var(--border);
  border-radius: 8px;
  padding: 8px 12px;
  margin-bottom: 6px;
  gap: 12px;
}
.cal-exercise-cue {
  font-size: 0.76em;
  color: var(--text-dim);
  margin-top: 2px;
  display: flex;
  align-items: center;
  gap: 4px;
}
.cal-exercises-list {
  display: flex;
  flex-direction: column;
  gap: 4px;
  max-height: 220px;
  overflow-y: auto;
  padding-right: 4px;
}

/* PIN Keypad & Lock Screen */
.pin-dot {
  width: 14px;
  height: 14px;
  border-radius: 50%;
  border: 2px solid rgba(255, 255, 255, 0.25);
  background: transparent;
  transition: all 0.15s ease;
}
.pin-dot.filled {
  background: var(--teal);
  border-color: var(--teal);
  box-shadow: 0 0 10px var(--teal-glow);
  transform: scale(1.1);
}

.btn-pin-key {
  height: 54px;
  border-radius: 14px;
  background: rgba(255, 255, 255, 0.04);
  border: 1px solid rgba(255, 255, 255, 0.09);
  color: #FFF;
  font-size: 1.35em;
  font-weight: 700;
  cursor: pointer;
  display: flex;
  align-items: center;
  justify-content: center;
  transition: all 0.15s ease;
  user-select: none;
  -webkit-tap-highlight-color: transparent;
}
.btn-pin-key:active {
  background: rgba(0, 229, 255, 0.15);
  border-color: var(--teal);
  transform: scale(0.95);
}
.btn-pin-key.pin-func {
  font-size: 0.9em;
  color: var(--text-muted);
}

.pin-shake {
  animation: pinShakeAnim 0.35s ease;
}
@keyframes pinShakeAnim {
  0%, 100% { transform: translateX(0); }
  20%, 60% { transform: translateX(-8px); }
  40%, 80% { transform: translateX(8px); }
}

/* Cloud Sync Header Status Dot */
.cloud-sync-status-dot {
  position: absolute;
  top: 5px;
  right: 5px;
  width: 7px;
  height: 7px;
  border-radius: 50%;
  background: var(--text-dim);
  transition: all 0.2s ease;
  pointer-events: none;
}
.cloud-sync-status-dot.synced {
  background: #10B981;
  box-shadow: 0 0 8px #10B981;
}
.cloud-sync-status-dot.syncing {
  background: var(--teal);
  box-shadow: 0 0 8px var(--teal);
  animation: pulse-sync-dot 1.2s infinite ease-in-out;
}
.cloud-sync-status-dot.pending {
  background: #F59E0B;
  box-shadow: 0 0 8px #F59E0B;
}
@keyframes pulse-sync-dot {
  0% { transform: scale(0.9); opacity: 0.8; }
  50% { transform: scale(1.3); opacity: 1; }
  100% { transform: scale(0.9); opacity: 0.8; }
}

/* =======================================================
   MOBILE PHONE UX & ERGONOMICS (Android PWA Responsive)
   ======================================================= */

/* Quick-Tap Chips */
.quick-chips-bar {
  display: flex;
  gap: 6px;
  overflow-x: auto;
  padding: 4px 0 8px 0;
  margin-bottom: 4px;
  -webkit-overflow-scrolling: touch;
  scrollbar-width: none;
}
.quick-chips-bar::-webkit-scrollbar {
  display: none;
}
.quick-chip {
  flex-shrink: 0;
  padding: 5px 10px;
  border-radius: 8px;
  background: rgba(255, 255, 255, 0.05);
  border: 1px solid rgba(255, 255, 255, 0.1);
  color: var(--text-muted);
  font-size: 0.76em;
  font-weight: 700;
  cursor: pointer;
  transition: all 0.15s ease;
  user-select: none;
  -webkit-tap-highlight-color: transparent;
}
.quick-chip:hover, .quick-chip:focus {
  background: rgba(0, 229, 255, 0.12);
  border-color: var(--teal);
  color: var(--teal);
}
.quick-chip:active {
  background: var(--teal);
  color: #000;
  transform: scale(0.95);
}

.sheet-drag-pill {
  display: none;
}

.mobile-bottom-nav {
  display: none;
}

@media (max-width: 768px) {
  /* Safe viewport height & padding for Android system gestures */
  html, body {
    overflow-x: hidden !important;
    width: 100% !important;
    max-width: 100vw !important;
  }
  body {
    padding-bottom: calc(68px + env(safe-area-inset-bottom, 14px)) !important;
  }

  .view-viewport,
  .view-screen {
    overflow-x: hidden !important;
    width: 100% !important;
    max-width: 100vw !important;
    box-sizing: border-box !important;
  }

  /* Compact Mobile Top App Bar */
  .master-header {
    padding: 10px 14px !important;
    gap: 10px !important;
    width: 100% !important;
    box-sizing: border-box !important;
  }
  .header-nav-tabs,
  #nav-breadcrumb,
  #header-actions {
    display: none !important;
  }
  .brand-logo {
    font-size: 1.05em !important;
    gap: 6px !important;
  }
  .header-right-actions {
    gap: 6px !important;
  }
  .btn-header {
    padding: 6px 10px !important;
    font-size: 0.82em !important;
    box-sizing: border-box !important;
  }

  /* Native Mobile Bottom Navigation Bar */
  .mobile-bottom-nav {
    display: flex !important;
    position: fixed;
    bottom: 0;
    left: 0;
    right: 0;
    height: calc(58px + env(safe-area-inset-bottom, 12px));
    padding-bottom: env(safe-area-inset-bottom, 12px);
    background: rgba(9, 13, 22, 0.95);
    backdrop-filter: blur(20px);
    -webkit-backdrop-filter: blur(20px);
    border-top: 1px solid rgba(255, 255, 255, 0.08);
    z-index: 999;
    align-items: center;
    justify-content: space-around;
    box-sizing: border-box;
    box-shadow: 0 -4px 25px rgba(0, 0, 0, 0.6);
  }

  .mobile-nav-item {
    flex: 1;
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    gap: 3px;
    height: 100%;
    background: transparent;
    border: none;
    color: var(--text-muted);
    cursor: pointer;
    font-size: 0.72em;
    font-weight: 700;
    transition: all 0.15s ease;
    user-select: none;
    -webkit-tap-highlight-color: transparent;
    padding: 4px 0;
  }

  .mobile-nav-item .mobile-nav-icon {
    display: flex;
    align-items: center;
    justify-content: center;
    width: 24px;
    height: 24px;
    transition: transform 0.15s ease;
  }

  .mobile-nav-item .mobile-nav-icon svg {
    width: 22px;
    height: 22px;
    fill: currentColor;
  }

  .mobile-nav-item.active {
    color: var(--teal);
  }

  .mobile-nav-item.active .mobile-nav-icon {
    transform: translateY(-2px);
    filter: drop-shadow(0 0 8px var(--teal-glow));
  }

  .mobile-nav-item:active {
    transform: scale(0.92);
  }

  /* Bottom-Sheet Style Modals */
  .modal-overlay {
    align-items: flex-end !important;
    padding: 0 !important;
  }

  .modal-window {
    width: 100% !important;
    max-width: 100% !important;
    max-height: 88vh !important;
    border-radius: 22px 22px 0 0 !important;
    margin: 0 !important;
    transform: translateY(100%) !important;
    transition: transform 0.26s cubic-bezier(0.16, 1, 0.3, 1) !important;
    padding-bottom: calc(20px + env(safe-area-inset-bottom, 16px)) !important;
    border-bottom: none !important;
    box-shadow: 0 -10px 40px rgba(0, 0, 0, 0.7) !important;
    overflow-y: auto !important;
  }

  .modal-overlay.show .modal-window {
    transform: translateY(0) !important;
  }

  .sheet-drag-pill {
    display: block !important;
    width: 38px;
    height: 4px;
    border-radius: 2px;
    background: rgba(255, 255, 255, 0.22);
    margin: 0 auto 12px auto;
  }

  .form-row-2, .form-row-3 {
    grid-template-columns: 1fr !important;
    gap: 10px !important;
  }

  /* Workout Builder Vertical Mobile Exercise Cards */
  .exercise-row {
    display: flex !important;
    flex-direction: column !important;
    align-items: stretch !important;
    padding: 14px !important;
    border-radius: 16px !important;
    gap: 10px !important;
    background: var(--bg-card) !important;
    border: 1px solid var(--border) !important;
    margin-bottom: 12px !important;
    width: 100% !important;
    box-sizing: border-box !important;
  }

  .exercise-row .row-left {
    display: flex !important;
    align-items: center !important;
    justify-content: flex-start !important;
    width: 100% !important;
    gap: 8px !important;
    flex-wrap: wrap !important;
  }

  .exercise-row .row-title {
    font-size: 1.02em !important;
    font-weight: 800 !important;
    color: var(--text-main) !important;
    white-space: normal !important;
    overflow: visible !important;
    word-break: break-word !important;
  }

  .exercise-row .row-pill-info {
    display: grid !important;
    grid-template-columns: repeat(3, 1fr) !important;
    gap: 8px !important;
    width: 100% !important;
    margin: 2px 0 !important;
  }

  .exercise-row .row-badge {
    display: flex !important;
    flex-direction: column !important;
    align-items: center !important;
    justify-content: center !important;
    padding: 8px 6px !important;
    border-radius: 10px !important;
    background: rgba(255, 255, 255, 0.03) !important;
    border: 1px solid rgba(255, 255, 255, 0.06) !important;
    font-size: 0.78em !important;
    font-weight: 700 !important;
    color: var(--text-main) !important;
    gap: 4px !important;
    text-align: center !important;
  }

  .exercise-row .row-badge svg {
    width: 16px !important;
    height: 16px !important;
    color: var(--teal) !important;
  }

  .exercise-row .row-actions {
    display: flex !important;
    justify-content: flex-end !important;
    align-items: center !important;
    gap: 8px !important;
    width: 100% !important;
    padding-top: 8px !important;
    border-top: 1px solid rgba(255, 255, 255, 0.05) !important;
  }

  .exercise-row .btn-row-action {
    height: 36px !important;
    min-width: 36px !important;
    padding: 0 10px !important;
    border-radius: 8px !important;
  }

  .exercise-row .btn-row-action.edit-btn {
    flex: 1 !important;
    height: 38px !important;
    font-size: 0.84em !important;
    justify-content: center !important;
  }

  /* Coach cue inside mobile card */
  .row-cue-banner {
    background: rgba(167, 139, 250, 0.08);
    border-left: 3px solid #A78BFA;
    border-radius: 0 8px 8px 0;
    padding: 6px 10px;
    font-size: 0.78em;
    font-style: italic;
    color: #CBD5E1;
    line-height: 1.35;
  }

  /* Mobile Content Padding */
  .hub-content,
  .athlete-detail-content,
  .builder-content {
    padding: 12px 14px 80px 14px !important;
    width: 100% !important;
    max-width: 100% !important;
    box-sizing: border-box !important;
  }

  .athletes-grid,
  .templates-grid {
    grid-template-columns: 1fr !important;
    gap: 12px !important;
    width: 100% !important;
  }
  .athlete-roster-card {
    padding: 14px !important;
    width: 100% !important;
    box-sizing: border-box !important;
  }

  /* Hub Header Bar Mobile Stacking */
  .hub-header-bar {
    flex-direction: column !important;
    align-items: stretch !important;
    padding: 14px !important;
    margin-bottom: 14px !important;
    gap: 12px !important;
    width: 100% !important;
    box-sizing: border-box !important;
  }
  .hub-header-info {
    width: 100% !important;
  }
  .hub-header-actions {
    display: flex !important;
    gap: 8px !important;
    width: 100% !important;
    box-sizing: border-box !important;
  }
  .hub-search-input,
  #athlete-search-input,
  #template-search-input {
    flex: 1 !important;
    width: 100% !important;
    min-width: 0 !important;
    box-sizing: border-box !important;
    padding: 8px 10px !important;
  }
  .hub-header-actions .btn-header-primary {
    flex-shrink: 0 !important;
    padding: 0 12px !important;
    font-size: 0.82em !important;
    white-space: nowrap !important;
  }

  /* Athlete Profile Header Mobile */
  .athlete-profile-header {
    flex-direction: column !important;
    align-items: stretch !important;
    gap: 12px !important;
    margin-bottom: 14px !important;
    width: 100% !important;
    box-sizing: border-box !important;
  }
  .athlete-profile-info {
    display: flex !important;
    align-items: center !important;
    gap: 12px !important;
    width: 100% !important;
  }
  .athlete-profile-actions {
    display: grid !important;
    grid-template-columns: 1fr 1fr 1fr !important;
    gap: 6px !important;
    width: 100% !important;
    box-sizing: border-box !important;
  }
  .athlete-profile-actions .btn-header {
    justify-content: center !important;
    padding: 0 2px !important;
    font-size: 0.74em !important;
    height: 38px !important;
    white-space: nowrap !important;
    overflow: hidden !important;
    text-overflow: ellipsis !important;
  }

  /* Athlete Notes & Comments Textarea */
  .athlete-comments-card {
    padding: 14px !important;
    margin-bottom: 14px !important;
    width: 100% !important;
    box-sizing: border-box !important;
  }
  .comments-textarea {
    box-sizing: border-box !important;
    width: 100% !important;
    min-height: 80px !important;
  }

  /* Sessions Manager Header Bar */
  .sessions-manager-bar {
    margin-bottom: 12px !important;
    display: flex !important;
    justify-content: space-between !important;
    align-items: center !important;
    gap: 8px !important;
    width: 100% !important;
    box-sizing: border-box !important;
  }
  .sessions-manager-title {
    font-size: 1.05em !important;
    min-width: 0 !important;
  }
  .sessions-manager-bar .btn-header-primary {
    font-size: 0.78em !important;
    padding: 0 10px !important;
    flex-shrink: 0 !important;
    white-space: nowrap !important;
  }

  /* Mobile Session Overview Cards (Zero Overlap & Touch-Friendly) */
  .session-overview-card {
    flex-direction: column !important;
    align-items: stretch !important;
    padding: 14px !important;
    gap: 10px !important;
    width: 100% !important;
    box-sizing: border-box !important;
  }
  .soc-left {
    width: 100% !important;
    flex-direction: column !important;
    align-items: flex-start !important;
    gap: 8px !important;
    min-width: 0 !important;
  }
  .soc-header-row {
    display: flex !important;
    align-items: center !important;
    gap: 8px !important;
    width: 100% !important;
    flex-wrap: wrap !important;
  }
  .soc-title {
    flex: 1 !important;
    min-width: 100px !important;
    font-size: 1.02em !important;
    font-weight: 800 !important;
    white-space: normal !important;
    word-break: break-word !important;
    margin-bottom: 0 !important;
  }
  .soc-type-badge {
    margin-left: auto !important;
  }
  .soc-badges {
    display: flex !important;
    gap: 6px !important;
    flex-wrap: wrap !important;
  }
  .soc-preview-exercises {
    font-size: 0.76em !important;
    color: var(--text-dim) !important;
    white-space: normal !important;
    line-height: 1.35 !important;
    word-break: break-word !important;
  }
  .soc-actions {
    display: flex !important;
    flex-direction: column !important;
    gap: 8px !important;
    width: 100% !important;
    padding-top: 10px !important;
    border-top: 1px solid rgba(255, 255, 255, 0.06) !important;
  }
  .soc-btn-edit {
    width: 100% !important;
    height: 42px !important;
    font-size: 0.88em !important;
    font-weight: 800 !important;
    justify-content: center !important;
    color: var(--teal) !important;
    background: rgba(0, 229, 255, 0.08) !important;
    border-color: rgba(0, 229, 255, 0.35) !important;
  }
  .soc-secondary-actions {
    display: grid !important;
    grid-template-columns: 1fr 1fr 1fr 40px !important;
    gap: 6px !important;
    width: 100% !important;
  }
  .soc-btn-sub {
    height: 36px !important;
    padding: 0 4px !important;
    font-size: 0.74em !important;
    justify-content: center !important;
  }
  .soc-btn-del {
    height: 36px !important;
    width: 40px !important;
    padding: 0 !important;
    justify-content: center !important;
  }

  /* Athlete History Timeline Mobile */
  .athlete-history-card {
    padding: 14px !important;
    width: 100% !important;
    box-sizing: border-box !important;
  }
  .history-card-header {
    flex-wrap: wrap !important;
    gap: 8px !important;
    width: 100% !important;
  }
  .history-card-header > div:first-child {
    font-size: 0.95em !important;
  }
  .history-item-row {
    flex-direction: column !important;
    align-items: stretch !important;
    gap: 8px !important;
    width: 100% !important;
    box-sizing: border-box !important;
  }
  .history-item-row > div:first-child {
    width: 100% !important;
    min-width: 0 !important;
  }
  .history-item-summary {
    word-break: break-word !important;
    overflow-wrap: break-word !important;
    display: inline !important;
  }

  /* Workout Builder Toolbar & Add Bar Mobile */
  .builder-toolbar {
    flex-direction: column !important;
    align-items: stretch !important;
    gap: 8px !important;
    padding: 10px 12px !important;
    width: 100% !important;
    box-sizing: border-box !important;
  }
  .builder-tb-top {
    display: flex !important;
    justify-content: space-between !important;
    align-items: center !important;
    width: 100% !important;
  }
  .builder-tb-mid {
    display: flex !important;
    gap: 8px !important;
    width: 100% !important;
  }
  .builder-tb-mid .builder-session-input {
    flex: 1 !important;
    width: 100% !important;
    min-width: 0 !important;
    box-sizing: border-box !important;
  }
  .builder-tb-mid select {
    width: 120px !important;
    flex-shrink: 0 !important;
  }
  .builder-tb-bottom {
    display: grid !important;
    grid-template-columns: 1fr 1fr !important;
    gap: 8px !important;
    width: 100% !important;
  }
  .builder-tb-bottom .btn-header {
    justify-content: center !important;
    font-size: 0.78em !important;
  }
  .add-bar {
    flex-direction: column !important;
    gap: 8px !important;
    width: 100% !important;
  }
  .btn-add-action {
    width: 100% !important;
    padding: 12px !important;
    font-size: 0.84em !important;
  }

  /* Calendar Toolbar & Grid Mobile (Zero Overflow & 7-Col minmax) */
  .calendar-container {
    width: 100% !important;
    max-width: 100% !important;
    box-sizing: border-box !important;
  }
  .calendar-controls-bar {
    flex-direction: column !important;
    align-items: stretch !important;
    gap: 10px !important;
    padding: 12px !important;
    width: 100% !important;
    box-sizing: border-box !important;
  }
  .cal-nav-group {
    display: flex !important;
    justify-content: space-between !important;
    align-items: center !important;
    width: 100% !important;
  }
  .cal-month-title {
    min-width: 0 !important;
    font-size: 1em !important;
  }
  .cal-filter-group {
    display: flex !important;
    justify-content: space-between !important;
    align-items: center !important;
    width: 100% !important;
    gap: 8px !important;
  }
  .cal-filter-group select {
    flex: 1 !important;
    min-width: 0 !important;
  }
  .calendar-weekdays-grid {
    display: grid !important;
    grid-template-columns: repeat(7, minmax(0, 1fr)) !important;
    width: 100% !important;
    font-size: 0.65em !important;
    padding: 6px 0 !important;
  }
  .calendar-days-grid {
    display: grid !important;
    grid-template-columns: repeat(7, minmax(0, 1fr)) !important;
    width: 100% !important;
  }
  .calendar-day-cell {
    min-height: 56px !important;
    padding: 2px !important;
    min-width: 0 !important;
    overflow: hidden !important;
    box-sizing: border-box !important;
  }
  .cal-day-header {
    margin-bottom: 2px !important;
  }
  .cal-day-number {
    font-size: 0.68em !important;
  }
  .cal-add-log-btn {
    display: none !important;
  }
  .cal-badges-container {
    width: 100% !important;
    min-width: 0 !important;
    gap: 2px !important;
  }
  .cal-session-badge {
    font-size: 0.6em !important;
    padding: 2px 3px !important;
    width: 100% !important;
    min-width: 0 !important;
    box-sizing: border-box !important;
    overflow: hidden !important;
    display: block !important;
  }
  .cal-session-badge > span {
    width: 100% !important;
    min-width: 0 !important;
  }
  .cal-session-badge .cal-badge-ath-avatar {
    display: none !important;
  }


  /* History Card and Calendar Modal Mobile Enhancements */
  .cal-assign-grid {
    grid-template-columns: 1fr !important;
    gap: 10px !important;
  }
  .cal-assign-submit-btn {
    width: 100% !important;
    height: 42px !important;
    font-size: 0.88em !important;
    font-weight: 800 !important;
    justify-content: center !important;
  }
  .status-pill-group {
    display: grid !important;
    grid-template-columns: repeat(3, 1fr) !important;
    gap: 6px !important;
  }
  .status-pill-btn {
    justify-content: center !important;
    text-align: center !important;
    font-size: 0.78em !important;
    padding: 6px 2px !important;
  }

}

.plan-day-row {
  align-items: flex-start !important;
  padding: 14px 16px !important;
}
.plan-day-row .row-left {
  flex: 1 !important;
  display: flex !important;
  align-items: flex-start !important;
  gap: 12px !important;
}
.plan-day-badge {
  font-size: 0.76em !important;
  font-weight: 800 !important;
  text-transform: uppercase !important;
  letter-spacing: 0.04em !important;
  color: var(--purple) !important;
  background: rgba(167, 139, 250, 0.15) !important;
  padding: 3px 8px !important;
  border-radius: 6px !important;
}
.plan-day-focus {
  font-size: 1.05em !important;
  font-weight: 700 !important;
  color: var(--text-main) !important;
}
.plan-day-details {
  font-size: 0.85em !important;
  color: var(--text-muted) !important;
  line-height: 1.5 !important;
  white-space: pre-line !important;
  margin-top: 6px !important;
  background: rgba(0, 0, 0, 0.22) !important;
  border-radius: 8px !important;
  padding: 8px 12px !important;
  border-left: 3px solid var(--purple) !important;
}


"""

STUDIO_HTML = """
<!-- Master Top Navigation Header -->
<div class="master-header">
  <div class="brand-group">
    <div class="brand-logo" onclick="navigateTo('hub')">
      {MUI_ICONS['bolt']}
      <span>Coach Studio</span>
    </div>

    <!-- Master Navigation Tabs -->
    <div class="header-nav-tabs">
      <button class="header-tab-btn active" id="tab-nav-athletes" onclick="navigateTo('hub')">
        {MUI_ICONS['athletes']} Athletes
      </button>
      <button class="header-tab-btn" id="tab-nav-calendar" onclick="navigateTo('calendar')">
        {MUI_ICONS['calendar']} Calendar
      </button>
      <button class="header-tab-btn" id="tab-nav-templates" onclick="navigateTo('templates')">
        {MUI_ICONS['templates']} Session Templates
      </button>
    </div>

    <div class="nav-breadcrumb" id="nav-breadcrumb">
      <!-- Breadcrumb buttons injected dynamically -->
    </div>
  </div>

  <div class="header-right-actions">
    <div id="header-actions" style="display:flex; align-items:center; gap:8px;">
      <!-- Header buttons injected dynamically -->
    </div>

    <!-- 24/7 CLOUD SYNC QUICK BUTTON -->
    <button class="btn-header" id="btn-header-cloud-sync" onclick="openCloudSyncModal()" title="24/7 Cloud Sync" style="position:relative;">
      {MUI_ICONS['cloud_sync']}
      <span class="cloud-sync-status-dot" id="cloud-sync-dot"></span>
    </button>

    <!-- STUDIO TOOLS DROPDOWN MENU -->
    <div style="position:relative; display:inline-block;">
      <button class="btn-header" id="btn-studio-tools-menu" onclick="toggleStudioToolsMenu(event)" title="Studio Tools & Settings">
        <svg class="mui-icon" viewBox="0 0 24 24"><path d="M12 8c1.1 0 2-.9 2-2s-.9-2-2-2-2 .9-2 2 .9 2 2 2zm0 2c-1.1 0-2 .9-2 2s.9 2 2 2 2-.9 2-2-.9-2-2-2zm0 6c-1.1 0-2 .9-2 2s.9 2 2 2 2-.9 2-2-.9-2-2-2z"/></svg>
      </button>
      <div class="custom-select-menu" id="studio-tools-menu" style="right:0; left:auto; width:220px;">
        <div class="custom-select-item" onclick="navigateTo('calendar'); closeAllCustomSelects();">
          <span style="display:flex; align-items:center; gap:8px;">{MUI_ICONS['calendar']} Training Calendar</span>
        </div>
        <div class="custom-select-item" onclick="openImportRecapModal(); closeAllCustomSelects();">
          <span style="display:flex; align-items:center; gap:8px;">{MUI_ICONS['copy']} Import Workout Recap</span>
        </div>
        <div class="custom-select-item" onclick="openCloudSyncModal(); closeAllCustomSelects();">
          <span style="display:flex; align-items:center; gap:8px;">{MUI_ICONS['cloud']} 24/7 Cloud Sync</span>
        </div>
        <div class="custom-select-item" onclick="lockStudioNow(); closeAllCustomSelects();">
          <span style="display:flex; align-items:center; gap:8px;">{MUI_ICONS['lock']} Lock Studio Now</span>
        </div>
        <div class="custom-select-item" onclick="openPasscodeSettingsModal(); closeAllCustomSelects();">
          <span style="display:flex; align-items:center; gap:8px;">{MUI_ICONS['shield']} Passcode Security</span>
        </div>
        <div class="custom-select-item" onclick="openLibraryModal(); closeAllCustomSelects();">
          <span style="display:flex; align-items:center; gap:8px;">{MUI_ICONS['library']} Exercise Library</span>
        </div>
        <div class="custom-select-item" onclick="openBackupModal(); closeAllCustomSelects();">
          <span style="display:flex; align-items:center; gap:8px;">{MUI_ICONS['backup']} Backup & Sync (.json)</span>
        </div>
        <div class="custom-select-item" onclick="downloadCsvTemplate(); closeAllCustomSelects();">
          <span style="display:flex; align-items:center; gap:8px;">{MUI_ICONS['download']} Download CSV Template</span>
        </div>
        <div class="custom-select-item" onclick="navigateTo('templates'); closeAllCustomSelects();">
          <span style="display:flex; align-items:center; gap:8px;">{MUI_ICONS['templates']} Session Templates</span>
        </div>
        <div class="custom-select-item" onclick="openNewAthleteModal(); closeAllCustomSelects();">
          <span style="display:flex; align-items:center; gap:8px;">{MUI_ICONS['add']} Create Athlete</span>
        </div>
      </div>
    </div>
  </div>
</div>

<!-- Viewport with 3 Dedicated Views -->
<div class="view-viewport">

  <!-- VIEW 1: ATHLETES ROSTER (HUB) -->
  <div class="view-screen active" id="view-hub">
    <div class="hub-content">
      <div class="hub-header-bar">
        <div class="hub-header-info">
          <div style="font-size:1.3em; font-weight:800; color:var(--text-main); display:flex; align-items:center; gap:10px;">
            <span>{MUI_ICONS['athletes']} Athlete Directory</span>
            <span class="athlete-badge-count" id="athlete-count-badge">2 Athletes</span>
          </div>
          <div style="font-size:0.84em; color:var(--text-muted); margin-top:4px;">Select an athlete to view their sessions, manage comments, or build programs.</div>
        </div>
        <div class="hub-header-actions">
          <input type="text" id="athlete-search-input" class="form-input hub-search-input" placeholder="Search athletes..." oninput="renderAthletesHub()">
          <button class="btn-header btn-header-primary" onclick="openNewAthleteModal()">
            {MUI_ICONS['add']} New Athlete
          </button>
        </div>
      </div>

      <div class="athletes-grid" id="athletes-grid">
        <!-- Rendered dynamically -->
      </div>
    </div>
  </div>

  <!-- VIEW 2: ATHLETE DETAIL / HIGH-LEVEL SESSIONS VIEW -->
  <div class="view-screen" id="view-athlete">
    <div class="athlete-detail-content">
      
      <!-- Athlete Header Bar -->
      <div class="athlete-profile-header">
        <div class="athlete-profile-info">
          <div class="athlete-avatar-circle" id="detail-athlete-avatar">JS</div>
          <div>
            <h2 style="font-size:1.45em; font-weight:800; color:var(--text-main);" id="detail-athlete-name">Jada Said</h2>
            <div style="font-size:0.82em; color:var(--text-muted); margin-top:2px;" id="detail-athlete-sub">Athlete Program Overview</div>
          </div>
        </div>

        <div class="athlete-profile-actions">
          <button class="btn-header" onclick="openPreviewModal()">
            {MUI_ICONS['phone']} Preview
          </button>
          <button class="btn-header" onclick="quickShareAthleteWhatsApp()" title="1-Tap WhatsApp Share" style="color:#10B981; border-color:rgba(16, 185, 129, 0.35);">
            {MUI_ICONS['whatsapp']} WhatsApp
          </button>
          <button class="btn-header btn-header-primary" onclick="openExportSessionsModal()">
            {MUI_ICONS['export']} Export App
          </button>
        </div>
      </div>

      <!-- Coach Comments & Notes Card -->
      <div class="athlete-comments-card">
        <div class="comments-header">
          <div class="comments-title">
            <span>{MUI_ICONS['comment']} Coach Comments & Athlete Notes</span>
          </div>
          <span style="font-size:0.75em; color:var(--text-dim);" id="comments-status">Auto-saved</span>
        </div>
        <textarea class="comments-textarea" id="athlete-comments-input" 
                  placeholder="Type athlete comments, injury history, recovery notes, focus areas..." 
                  oninput="saveCurrentAthleteComments(this.value)"></textarea>
      </div>

      <!-- High-Level Sessions Manager -->
      <div class="sessions-manager-bar">
        <div class="sessions-manager-title">
          <span>{MUI_ICONS['dumbbell']} Workout Sessions</span>
          <span class="athlete-badge-count" id="detail-sessions-count">10 Sessions</span>
        </div>
        <button class="btn-header btn-header-primary" onclick="createNewSessionPrompt()">
          {MUI_ICONS['add']} Add Session
        </button>
      </div>

      <!-- Sessions List -->
      <div class="sessions-overview-grid" id="sessions-overview-grid">
        <!-- Rendered dynamically -->
      </div>

      <div style="margin-top:20px; text-align:center;">
        <button class="btn-add-action" onclick="createNewSessionPrompt()">
          {MUI_ICONS['add']} Add Another Workout Session
        </button>
      </div>

      <!-- Recent Training Activity & Logs Timeline -->
      <div class="athlete-history-card">
        <div class="history-card-header">
          <div style="font-size:1.05em; font-weight:800; color:var(--text-main); display:flex; align-items:center; gap:8px;">
            {MUI_ICONS['history']} Recent Training Logs & Workout History
          </div>
          <button class="btn-header btn-header-primary" onclick="openCalendarDayModal(null, activeAthleteId)" style="font-size:0.78em; padding:6px 14px;">
            {MUI_ICONS['add']} Log Workout
          </button>
        </div>
        <div class="history-timeline-list" id="athlete-history-timeline">
          <!-- Rendered dynamically -->
        </div>
      </div>

    </div>
  </div>

  <!-- VIEW 3: WORKOUT BUILDER / EXERCISE EDITOR -->
  <div class="view-screen" id="view-builder">
    <div class="builder-content">
      
      <div class="builder-toolbar">
        <div class="builder-tb-top">
          <button class="btn-nav-crumb" onclick="navigateTo('athlete')" title="Back to All Sessions">
            {MUI_ICONS['arrow_back']} <span>Back to Sessions</span>
          </button>
          <button class="btn-header btn-header-primary" onclick="saveActiveAthleteRoutine()">
            {MUI_ICONS['save']} Save Routine
          </button>
        </div>

        <div class="builder-tb-mid">
          <input type="text" id="builder-session-title" class="builder-session-input" placeholder="Session Title" onchange="updateActiveSessionTitle(this.value)">
          <select id="builder-session-type" class="form-input" style="width:130px; font-weight:700;" onchange="updateActiveSessionType(this.value)">
            <option value="exercises">Exercise List</option>
            <option value="table">Weekly Plan</option>
          </select>
        </div>

        <div class="builder-tb-bottom">
          <button class="btn-header" onclick="openCopySessionModal(activeSessionIndex)">
            Copy to Athlete...
          </button>
          <button class="btn-header" onclick="openPreviewModal()">
            {MUI_ICONS['phone']} Preview
          </button>
        </div>
      </div>

      <!-- Compact Exercise Rows Container -->
      <div id="exercises-container">
        <!-- Injected dynamically -->
      </div>

      <!-- Add Actions -->
      <div class="add-bar" id="builder-add-bar">
        <button class="btn-add-action" onclick="openExerciseModal(null)">
          {MUI_ICONS['add']} Add Blank Exercise
        </button>
        <button class="btn-add-action lib-btn" onclick="openLibraryModal()">
          {MUI_ICONS['library']} Pick from Exercise Library
        </button>
        <button class="btn-add-action" onclick="openBatchImportModal()" style="border-color:rgba(167,139,250,0.4); color:var(--purple);">
          {MUI_ICONS['copy']} Paste from Excel / CSV
        </button>
      </div>

    </div>
  </div>

  <!-- VIEW 5: ATHLETE TRAINING CALENDAR -->
  <div class="view-screen" id="view-calendar">
    <div class="hub-content">
      <div class="hub-header-bar">
        <div class="hub-header-info">
          <div style="font-size:1.3em; font-weight:800; color:var(--text-main); display:flex; align-items:center; gap:10px;">
            <span>{MUI_ICONS['calendar']} Athlete Training Calendar</span>
            <span class="athlete-badge-count" id="calendar-total-logs-badge">0 Logs</span>
          </div>
          <div style="font-size:0.84em; color:var(--text-muted); margin-top:4px;">
            Track when each athlete performed their sessions, log completed workouts, and import athlete recaps.
          </div>
        </div>
        <div class="hub-header-actions">
          <button class="btn-header" onclick="openImportRecapModal()">
            {MUI_ICONS['copy']} Import Recap
          </button>
          <button class="btn-header btn-header-primary" onclick="openCalendarDayModal()">
            {MUI_ICONS['add']} Log Workout
          </button>
        </div>
      </div>

      <!-- Calendar Controls & Stats Toolbar -->
      <div class="calendar-controls-bar">
        <!-- Month Switcher -->
        <div class="cal-nav-group">
          <button class="btn-header" onclick="navCalendarMonth(-1)" title="Previous Month">
            <svg class="mui-icon" viewBox="0 0 24 24"><path d="M15.41 7.41L14 6l-6 6 6 6 1.41-1.41L10.83 12z"/></svg>
          </button>
          <div class="cal-month-title" id="cal-month-title">September 2026</div>
          <button class="btn-header" onclick="navCalendarMonth(1)" title="Next Month">
            <svg class="mui-icon" viewBox="0 0 24 24"><path d="M10 6L8.59 7.41 13.17 12l-4.58 4.59L10 18l6-6z"/></svg>
          </button>
          <button class="btn-header" onclick="goToCalendarToday()" style="font-size:0.82em; padding:6px 12px; margin-left:4px;">
            Today
          </button>
        </div>

        <!-- Filter & Stats -->
        <div class="cal-filter-group">
          <label style="font-size:0.82em; color:var(--text-muted); font-weight:600;">Filter Athlete:</label>
          <select id="cal-athlete-filter" class="form-input" style="width:180px; padding:6px 10px; font-size:0.85em;" onchange="changeCalendarAthleteFilter(this.value)">
            <option value="all">All Athletes</option>
          </select>
          <div class="cal-stats-pill" id="cal-month-stats">
            0 Completed
          </div>
        </div>
      </div>

      <!-- Calendar 7-Column Grid Container -->
      <div class="calendar-container">
        <div class="calendar-weekdays-grid">
          <div>Mon</div>
          <div>Tue</div>
          <div>Wed</div>
          <div>Thu</div>
          <div>Fri</div>
          <div>Sat</div>
          <div>Sun</div>
        </div>
        <div class="calendar-days-grid" id="calendar-days-grid">
          <!-- Populated dynamically -->
        </div>
      </div>
    </div>
  </div>

  <!-- VIEW 4: MASTER SESSION TEMPLATES -->
  <div class="view-screen" id="view-templates">
    <div class="hub-content">
      <div class="hub-header-bar">
        <div class="hub-header-info">
          <div style="font-size:1.3em; font-weight:800; color:var(--text-main); display:flex; align-items:center; gap:10px;">
            <span>{MUI_ICONS['templates']} Master Session Templates</span>
            <span class="athlete-badge-count" id="template-count-badge">4 Templates</span>
          </div>
          <div style="font-size:0.84em; color:var(--text-muted); margin-top:4px;">Master workout blueprints you can seamlessly assign to any athlete with custom names.</div>
        </div>
        <div class="hub-header-actions">
          <input type="text" id="template-search-input" class="form-input hub-search-input" placeholder="Search templates..." oninput="renderTemplatesView()">
          <button class="btn-header btn-header-primary" onclick="openCreateTemplateModal()">
            {MUI_ICONS['add']} New Template
          </button>
        </div>
      </div>

      <div class="templates-grid" id="templates-grid" style="display:grid; grid-template-columns: repeat(auto-fill, minmax(330px, 1fr)); gap: 20px;">
        <!-- Rendered dynamically -->
      </div>
    </div>
  </div>

</div>

<!-- IN-APP TOAST NOTIFICATION -->
<div id="toast" class="toast-notification">
  <span id="toast-icon"></span>
  <span id="toast-message">Notification message</span>
</div>

<!-- IN-APP SLEEK CONFIRM MODAL -->
<div class="modal-overlay" id="custom-confirm-modal">
  <div class="modal-window" style="max-width:420px; text-align:center;">
    <div style="font-size:2.2em; margin-bottom:10px;" id="confirm-icon"></div>
    <div style="font-size:1.2em; font-weight:800; margin-bottom:8px; color:var(--text-main);" id="confirm-title">Confirm Action</div>
    <div style="font-size:0.88em; color:var(--text-muted); line-height:1.5; margin-bottom:24px;" id="confirm-message">Are you sure?</div>
    <div style="display:flex; justify-content:center; gap:12px;">
      <button class="btn-header" onclick="closeConfirmModal(false)">Cancel</button>
      <button class="btn-header btn-header-primary" id="confirm-ok-btn" onclick="closeConfirmModal(true)">Confirm</button>
    </div>
  </div>
</div>

<!-- IN-APP SLEEK PROMPT MODAL -->
<div class="modal-overlay" id="custom-prompt-modal">
  <div class="modal-window" style="max-width:460px;">
    <div class="modal-header-bar">
      <div class="modal-title" id="prompt-modal-title">Enter Value</div>
      <button class="modal-close-btn" onclick="closePromptModal(false)"><svg class="mui-icon" viewBox="0 0 24 24" style="width:18px;height:18px;"><path d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z"/></svg></button>
    </div>
    <div class="form-group">
      <label class="form-label" id="prompt-modal-label">Session Name</label>
      <input type="text" id="prompt-modal-input" class="form-input">
    </div>
    <div class="modal-footer-bar">
      <button class="btn-header" onclick="closePromptModal(false)">Cancel</button>
      <button class="btn-header btn-header-primary" onclick="closePromptModal(true)">Done</button>
    </div>
  </div>
</div>

<!-- MODAL: NEW ATHLETE -->
<div class="modal-overlay" id="new-athlete-modal">
  <div class="modal-window">
    <div class="modal-header-bar">
      <div class="modal-title">Create New Athlete Profile</div>
      <button class="modal-close-btn" onclick="closeModal('new-athlete-modal')"><svg class="mui-icon" viewBox="0 0 24 24" style="width:18px;height:18px;"><path d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z"/></svg></button>
    </div>
    
    <div class="form-group">
      <label class="form-label">Athlete Full Name *</label>
      <input type="text" id="new-ath-name" class="form-input" placeholder="e.g. Jordan Miller">
    </div>

    <div class="form-group">
      <label class="form-label">Initial Workout Program</label>
      <select id="new-ath-template" class="form-input">
        <option value="jada">Clone Jada Said's Routine (10 Sessions)</option>
        <option value="marcus">Clone Marcus Vance's Routine (2 Sessions)</option>
        <option value="blank">Start with 1 Blank Session</option>
      </select>
    </div>

    <div class="form-group">
      <label class="form-label">Coach Comments & Notes</label>
      <textarea id="new-ath-comments" class="form-textarea" placeholder="Injury history, specific training considerations..."></textarea>
    </div>

    <div class="modal-footer-bar">
      <button class="btn-header" onclick="closeModal('new-athlete-modal')">Cancel</button>
      <button class="btn-header btn-header-primary" onclick="confirmCreateAthlete()">Create Athlete</button>
    </div>
  </div>
</div>

<!-- MODAL: EXPORT SESSIONS SELECTION -->
<div class="modal-overlay" id="export-sessions-modal">
  <div class="modal-window" style="max-width:520px;">
    <div class="modal-header-bar">
      <div class="modal-title" style="display:flex; align-items:center; gap:8px;">
        <span style="color:var(--teal); display:inline-flex;">{MUI_ICONS['export']}</span>
        <span>Export Standalone Program (.html)</span>
      </div>
      <button class="modal-close-btn" onclick="closeModal('export-sessions-modal')">{MUI_ICONS['close']}</button>
    </div>

    <div style="font-size:0.88em; color:var(--text-muted); margin-bottom:14px;">
      Choose which sessions to include in the standalone HTML file for <strong id="export-modal-athlete-name" style="color:var(--teal);">Athlete</strong>:
    </div>

    <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:10px;">
      <span style="font-size:0.78em; font-weight:700; text-transform:uppercase; letter-spacing:0.5px; color:var(--text-dim);">Available Sessions</span>
      <div style="display:flex; gap:8px;">
        <button type="button" class="btn-sm" style="font-size:0.75em; padding:4px 10px; border-radius:6px; background:rgba(255,255,255,0.06); color:var(--text-muted); border:1px solid rgba(255,255,255,0.1); cursor:pointer;" onclick="selectAllExportSessions(true)">Select All</button>
        <button type="button" class="btn-sm" style="font-size:0.75em; padding:4px 10px; border-radius:6px; background:rgba(255,255,255,0.06); color:var(--text-muted); border:1px solid rgba(255,255,255,0.1); cursor:pointer;" onclick="selectAllExportSessions(false)">Deselect All</button>
      </div>
    </div>

    <div id="export-sessions-list" class="export-sessions-container" style="max-height:260px; overflow-y:auto; display:flex; flex-direction:column; gap:8px; margin-bottom:16px; padding-right:4px;">
      <!-- Populated dynamically -->
    </div>

    <div class="form-group" style="margin-bottom:16px;">
      <label class="form-label" style="display:flex; justify-content:space-between;">
        <span>Export File Name</span>
        <span style="font-size:0.8em; color:var(--text-dim);">Self-contained .html</span>
      </label>
      <input type="text" id="export-filename-input" class="form-input" placeholder="Program.html" style="font-family:monospace; font-size:0.9em;">
    </div>

    <div style="font-size:0.78em; color:var(--text-dim); margin-bottom:16px; display:flex; align-items:center; gap:6px;">
      <svg class="mui-icon" viewBox="0 0 24 24" style="width:14px;height:14px;fill:var(--teal);flex-shrink:0;"><path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm1 15h-2v-6h2v6zm0-8h-2V7h2v2z"/></svg>
      <span>Includes offline workout tracker, plate calculator, rest timer & iOS compatibility.</span>
    </div>

    <div class="modal-footer-bar" style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:10px;">
      <button class="btn-header" onclick="closeModal('export-sessions-modal')">Cancel</button>
      <div style="display:flex; gap:8px; align-items:center;">
        <button class="btn-header" id="export-download-btn" onclick="confirmExportSelectedSessions('download')" title="Download HTML file">
          {MUI_ICONS['download']} Download (.html)
        </button>
        <button class="btn-header btn-header-primary" id="export-share-btn" onclick="confirmExportSelectedSessions('share')" title="Send directly via WhatsApp" style="background:linear-gradient(135deg, #10B981 0%, #059669 100%); border-color:#10B981; color:#FFF; font-weight:800;">
          {MUI_ICONS['whatsapp']} Send via WhatsApp
        </button>
      </div>
    </div>
  </div>
</div>

<!-- MODAL: 24/7 CLOUD SYNC -->
<div class="modal-overlay" id="cloud-sync-modal">
  <div class="modal-window" style="max-width:540px;">
    <div class="sheet-drag-pill"></div>
    <div class="modal-header-bar">
      <div class="modal-title">{MUI_ICONS['cloud_sync']} 24/7 Cloud Sync (Zero Cost)</div>
      <button class="modal-close-btn" onclick="closeModal('cloud-sync-modal')"><svg class="mui-icon" viewBox="0 0 24 24" style="width:18px;height:18px;"><path d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z"/></svg></button>
    </div>

    <div style="margin:14px 0; display:flex; flex-direction:column; gap:14px;">
      <div style="font-size:0.84em; color:var(--text-muted); line-height:1.5;">
        Sync athletes, routines, and calendar history securely between your PC and Android phone using your free private GitHub Gist. Zero hosting fees, available 24/7 even when your PC is turned off.
      </div>

      <!-- Sync Status Card -->
      <div style="background:var(--bg-card); border:1px solid var(--border); border-radius:12px; padding:14px; display:flex; justify-content:space-between; align-items:center;">
        <div>
          <div style="font-size:0.75em; text-transform:uppercase; letter-spacing:0.5px; color:var(--text-dim); font-weight:700;">Sync Status</div>
          <div id="cloud-sync-status-text" style="font-size:0.95em; font-weight:800; color:var(--text-main); margin-top:2px;">Not Configured</div>
          <div id="cloud-sync-last-time" style="font-size:0.75em; color:var(--text-muted); margin-top:2px;">Never synced</div>
        </div>
        <div style="display:flex; gap:8px;">
          <button type="button" class="btn-header btn-header-primary" id="btn-cloud-sync-now" onclick="triggerManualCloudSync()" style="font-size:0.8em; padding:6px 12px;">
            {MUI_ICONS['refresh']} Sync Now
          </button>
        </div>
      </div>

      <!-- Configuration Form -->
      <div style="background:var(--bg-card); border:1px solid var(--border); border-radius:12px; padding:14px; display:flex; flex-direction:column; gap:10px;">
        <div style="font-size:0.85em; font-weight:800; color:var(--text-main);">Cloud Sync Settings</div>

        <div class="form-group" style="margin:0;">
          <label class="form-label" style="font-size:0.78em;">
            GitHub Personal Access Token (PAT)
            <span style="font-size:0.75em; color:var(--text-dim); font-weight:400; float:right;">(Needs 'gist' permission)</span>
          </label>
          <input type="password" id="cloud-sync-token-input" class="form-input" placeholder="ghp_xxxxxxxxxxxxxxxxxxxx" style="font-family:monospace; font-size:0.85em; padding:8px 10px;">
        </div>

        <div class="form-group" style="margin:0;">
          <label class="form-label" style="font-size:0.78em;">
            Private Gist ID
            <span style="font-size:0.75em; color:var(--text-dim); font-weight:400; float:right;">(Leave empty to auto-create)</span>
          </label>
          <input type="text" id="cloud-sync-gist-id-input" class="form-input" placeholder="Auto-generated upon first sync" style="font-family:monospace; font-size:0.85em; padding:8px 10px;">
        </div>

        <div style="display:flex; justify-content:space-between; align-items:center; margin-top:4px;">
          <label style="display:flex; align-items:center; gap:8px; font-size:0.82em; color:var(--text-muted); cursor:pointer;">
            <input type="checkbox" id="cloud-sync-auto-checkbox" checked style="accent-color:var(--teal); width:16px; height:16px;">
            <span>Auto-sync when opening and editing</span>
          </label>
          <button type="button" class="btn-sm" onclick="saveCloudSyncSettingsFromUI()" style="background:rgba(0, 229, 255, 0.12); color:var(--teal); border:1px solid rgba(0, 229, 255, 0.3); padding:5px 12px; border-radius:6px; font-weight:700; cursor:pointer;">Save Settings</button>
        </div>
      </div>

      <!-- Quick Actions -->
      <div style="display:grid; grid-template-columns:1fr 1fr; gap:10px;">
        <button class="btn-header" onclick="forcePushToCloud()" style="justify-content:center; font-size:0.8em; padding:8px;">
          {MUI_ICONS['export']} Push Local to Cloud
        </button>
        <button class="btn-header" onclick="forcePullFromCloud()" style="justify-content:center; font-size:0.8em; padding:8px;">
          {MUI_ICONS['download']} Pull Cloud to Local
        </button>
      </div>

      <div style="font-size:0.75em; color:var(--text-dim); line-height:1.4;">
        Security Notice: Your token is stored only in this browser's local storage and is never sent to any third-party server. All sync requests go directly from your browser to GitHub's encrypted API.
      </div>
    </div>

    <div class="modal-footer-bar">
      <button class="btn-header" onclick="closeModal('cloud-sync-modal')">Close</button>
    </div>
  </div>
</div>

<!-- MODAL: PASSCODE SECURITY SETTINGS -->
<div class="modal-overlay" id="passcode-modal">
  <div class="modal-window" style="max-width:440px;">
    <div class="sheet-drag-pill"></div>
    <div class="modal-header-bar">
      <div class="modal-title">{MUI_ICONS['lock']} Master Passcode Security</div>
      <button class="modal-close-btn" onclick="closeModal('passcode-modal')"><svg class="mui-icon" viewBox="0 0 24 24" style="width:18px;height:18px;"><path d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z"/></svg></button>
    </div>

    <div style="margin:16px 0; display:flex; flex-direction:column; gap:14px;">
      <div style="font-size:0.84em; color:var(--text-muted); line-height:1.5;">
        Set a 4 to 6 digit PIN to protect Coach Studio from unauthorized access on this device.
      </div>

      <div style="background:var(--bg-card); border:1px solid var(--border); border-radius:12px; padding:14px; display:flex; flex-direction:column; gap:12px;">
        <div class="form-group" style="margin:0;">
          <label class="form-label" style="font-size:0.8em;">New PIN (4-6 digits)</label>
          <input type="password" id="new-pin-input" class="form-input" maxlength="6" inputmode="numeric" placeholder="Enter 4-6 digits" style="font-size:1.1em; letter-spacing:4px; text-align:center;">
        </div>

        <div class="form-group" style="margin:0;">
          <label class="form-label" style="font-size:0.8em;">Confirm PIN</label>
          <input type="password" id="confirm-pin-input" class="form-input" maxlength="6" inputmode="numeric" placeholder="Re-enter digits" style="font-size:1.1em; letter-spacing:4px; text-align:center;">
        </div>
      </div>

      <div id="passcode-status-msg" style="font-size:0.82em; color:var(--text-muted);"></div>
    </div>

    <div class="modal-footer-bar" style="display:flex; justify-content:space-between;">
      <button class="btn-header" id="btn-disable-pin" onclick="disablePinSecurity()" style="color:var(--rose); display:none;">Disable PIN</button>
      <div style="display:flex; gap:8px;">
        <button class="btn-header" onclick="closeModal('passcode-modal')">Cancel</button>
        <button class="btn-header btn-header-primary" onclick="saveNewPinSecurity()">Save PIN</button>
      </div>
    </div>
  </div>
</div>

<!-- PASSCODE LOCK SCREEN OVERLAY -->
<div id="passcode-lock-screen" style="position:fixed; inset:0; background:#070A10; z-index:99999; display:none; flex-direction:column; align-items:center; justify-content:center; padding:20px; box-sizing:border-box;">
  <div style="width:100%; max-width:340px; display:flex; flex-direction:column; align-items:center; text-align:center;">
    <div style="width:64px; height:64px; border-radius:18px; background:rgba(0, 229, 255, 0.08); border:1px solid rgba(0, 229, 255, 0.25); display:flex; align-items:center; justify-content:center; margin-bottom:16px; color:var(--teal);">
      {MUI_ICONS['shield']}
    </div>
    <div style="font-size:1.3em; font-weight:900; color:#FFF; margin-bottom:6px;">Coach Studio Locked</div>
    <div style="font-size:0.84em; color:var(--text-muted); margin-bottom:24px;">Enter your master PIN to access the studio</div>

    <!-- PIN Dots Display -->
    <div id="pin-dots-container" style="display:flex; gap:14px; margin-bottom:24px;">
      <span class="pin-dot" id="pdot-0"></span>
      <span class="pin-dot" id="pdot-1"></span>
      <span class="pin-dot" id="pdot-2"></span>
      <span class="pin-dot" id="pdot-3"></span>
      <span class="pin-dot" id="pdot-4"></span>
      <span class="pin-dot" id="pdot-5"></span>
    </div>

    <!-- Hidden Input for keyboard typing -->
    <input type="password" id="pin-keyboard-input" maxlength="6" inputmode="numeric" pattern="[0-9]*" style="position:absolute; opacity:0; pointer-events:none;" oninput="handlePinKeyInput(this.value)">

    <!-- Number Pad -->
    <div class="pin-keypad" id="pin-keypad" style="display:grid; grid-template-columns:repeat(3, 1fr); gap:12px; width:100%; max-width:280px; margin-bottom:20px;">
      <button type="button" class="btn-pin-key" onclick="pressPinDigit('1')">1</button>
      <button type="button" class="btn-pin-key" onclick="pressPinDigit('2')">2</button>
      <button type="button" class="btn-pin-key" onclick="pressPinDigit('3')">3</button>
      <button type="button" class="btn-pin-key" onclick="pressPinDigit('4')">4</button>
      <button type="button" class="btn-pin-key" onclick="pressPinDigit('5')">5</button>
      <button type="button" class="btn-pin-key" onclick="pressPinDigit('6')">6</button>
      <button type="button" class="btn-pin-key" onclick="pressPinDigit('7')">7</button>
      <button type="button" class="btn-pin-key" onclick="pressPinDigit('8')">8</button>
      <button type="button" class="btn-pin-key" onclick="pressPinDigit('9')">9</button>
      <button type="button" class="btn-pin-key pin-func" onclick="clearPinInput()">Clear</button>
      <button type="button" class="btn-pin-key" onclick="pressPinDigit('0')">0</button>
      <button type="button" class="btn-pin-key pin-func" onclick="deletePinDigit()">Del</button>
    </div>

    <div id="pin-error-text" style="color:#EF4444; font-size:0.82em; font-weight:700; min-height:20px;"></div>
  </div>
</div>

<!-- MODAL: COPY SESSION TO ATHLETE -->
<div class="modal-overlay" id="copy-session-modal">
  <div class="modal-window" style="max-width:460px;">
    <div class="modal-header-bar">
      <div class="modal-title">Copy Session to Athlete</div>
      <button class="modal-close-btn" onclick="closeModal('copy-session-modal')"><svg class="mui-icon" viewBox="0 0 24 24" style="width:18px;height:18px;"><path d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z"/></svg></button>
    </div>

    <div style="font-size:0.88em; color:var(--text-muted); margin-bottom:16px;">
      Duplicate <strong id="copy-session-name" style="color:var(--teal);">Session</strong> into another athlete's workout plan:
    </div>

    <div class="form-group">
      <label class="form-label">Select Destination Athlete</label>
      <select id="copy-destination-athlete" class="form-input">
        <!-- Rendered dynamically -->
      </select>
    </div>

    <div class="modal-footer-bar">
      <button class="btn-header" onclick="closeModal('copy-session-modal')">Cancel</button>
      <button class="btn-header btn-header-primary" onclick="confirmCopySessionToAthlete()">Copy Workout</button>
    </div>
  </div>
</div>

<!-- MODAL: EDIT DAY / SCHEDULE BLOCK -->
<div class="modal-overlay" id="day-modal">
  <div class="modal-window">
    <div class="sheet-drag-pill"></div>
    <div class="modal-header-bar">
      <div class="modal-title" id="day-modal-title">{MUI_ICONS['calendar']} Edit Schedule Day</div>
      <button class="modal-close-btn" onclick="closeModal('day-modal')"><svg class="mui-icon" viewBox="0 0 24 24" style="width:18px;height:18px;"><path d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z"/></svg></button>
    </div>

    <div class="form-group">
      <label class="form-label">Day Label / Name *</label>
      <div class="quick-chips-bar" style="margin-bottom:8px; display:flex; flex-wrap:wrap; gap:6px;">
        <button type="button" class="quick-chip" onclick="quickFillDayField('modal-day-name', 'Monday')">Monday</button>
        <button type="button" class="quick-chip" onclick="quickFillDayField('modal-day-name', 'Tuesday')">Tuesday</button>
        <button type="button" class="quick-chip" onclick="quickFillDayField('modal-day-name', 'Wednesday')">Wednesday</button>
        <button type="button" class="quick-chip" onclick="quickFillDayField('modal-day-name', 'Thursday')">Thursday</button>
        <button type="button" class="quick-chip" onclick="quickFillDayField('modal-day-name', 'Friday')">Friday</button>
        <button type="button" class="quick-chip" onclick="quickFillDayField('modal-day-name', 'Saturday')">Saturday</button>
        <button type="button" class="quick-chip" onclick="quickFillDayField('modal-day-name', 'Sunday')">Sunday</button>
        <button type="button" class="quick-chip" onclick="quickFillDayField('modal-day-name', 'Rest Day')">Rest Day</button>
      </div>
      <input type="text" id="modal-day-name" class="form-input" placeholder="e.g. Monday or Day 1 - Push">
    </div>

    <div class="form-group">
      <label class="form-label">Training Focus / Title</label>
      <div class="quick-chips-bar" style="margin-bottom:8px; display:flex; flex-wrap:wrap; gap:6px;">
        <button type="button" class="quick-chip" onclick="quickFillDayField('modal-day-focus', 'Upper Body Power')">Upper Body</button>
        <button type="button" class="quick-chip" onclick="quickFillDayField('modal-day-focus', 'Lower Body Strength')">Lower Body</button>
        <button type="button" class="quick-chip" onclick="quickFillDayField('modal-day-focus', 'Push (Chest/Shoulders/Triceps)')">Push</button>
        <button type="button" class="quick-chip" onclick="quickFillDayField('modal-day-focus', 'Pull (Back/Biceps)')">Pull</button>
        <button type="button" class="quick-chip" onclick="quickFillDayField('modal-day-focus', 'Legs & Core')">Legs</button>
        <button type="button" class="quick-chip" onclick="quickFillDayField('modal-day-focus', 'Conditioning & Mobility')">Conditioning</button>
        <button type="button" class="quick-chip" onclick="quickFillDayField('modal-day-focus', 'Full Recovery & Rest')">Recovery</button>
      </div>
      <input type="text" id="modal-day-focus" class="form-input" placeholder="e.g. Upper Body Hypertrophy or Active Recovery">
    </div>

    <div class="form-group">
      <label class="form-label">Workout Overview & Prescribed Routine Details</label>
      <textarea id="modal-day-details" class="form-textarea" style="min-height:110px; font-family:inherit; line-height:1.5;" placeholder="List exercises, sets, reps, or coaching instructions for this day"></textarea>
    </div>

    <div class="modal-footer-bar">
      <button class="btn-header" onclick="closeModal('day-modal')">Cancel</button>
      <button class="btn-header btn-header-primary" onclick="saveDayModal()">Save Day</button>
    </div>
  </div>
</div>

<!-- MODAL: EDIT EXERCISE -->
<div class="modal-overlay" id="exercise-modal">
  <div class="modal-window">
    <div class="sheet-drag-pill"></div>
    <div class="modal-header-bar">
      <div class="modal-title" id="exercise-modal-title">{MUI_ICONS['edit']} Edit Exercise Details</div>
      <button class="modal-close-btn" onclick="closeModal('exercise-modal')"><svg class="mui-icon" viewBox="0 0 24 24" style="width:18px;height:18px;"><path d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z"/></svg></button>
    </div>

    <div class="form-row-2">
      <div class="form-group">
        <label class="form-label">Superset / Group Tag (Optional)</label>
        <input type="text" id="modal-ex-superset" class="form-input" placeholder="e.g. A1, A2, B1, B2">
      </div>
      <div class="form-group">
        <label class="form-label">Category / Phase</label>
        <select id="modal-ex-category" class="form-input">
          <option value="Warm-Up">Warm-Up</option>
          <option value="Primary Strength">Primary Strength</option>
          <option value="Secondary Strength">Secondary Strength</option>
          <option value="Accessory">Accessory</option>
          <option value="Core & Stability">Core & Stability</option>
          <option value="Conditioning / Finisher">Conditioning / Finisher</option>
          <option value="Cooldown">Cooldown</option>
        </select>
      </div>
      <div class="form-group">
        <label class="form-label">Exercise Name *</label>
        <input type="text" id="modal-ex-name" class="form-input" placeholder="e.g. DB Bench Press">
      </div>
    </div>

    <div class="form-row-3">
      <div class="form-group">
        <label class="form-label">Sets & Reps</label>
        <div class="quick-chips-bar">
          <button type="button" class="quick-chip" onclick="quickFillExerciseField('modal-ex-details', '3 x 10')">3 x 10</button>
          <button type="button" class="quick-chip" onclick="quickFillExerciseField('modal-ex-details', '4 x 8')">4 x 8</button>
          <button type="button" class="quick-chip" onclick="quickFillExerciseField('modal-ex-details', '3 x 12')">3 x 12</button>
          <button type="button" class="quick-chip" onclick="quickFillExerciseField('modal-ex-details', '5 x 5')">5 x 5</button>
          <button type="button" class="quick-chip" onclick="quickFillExerciseField('modal-ex-details', '3 x 15')">3 x 15</button>
        </div>
        <input type="text" id="modal-ex-details" class="form-input" placeholder="3 x 10">
      </div>
      <div class="form-group">
        <label class="form-label">Target Weight</label>
        <div class="quick-chips-bar">
          <button type="button" class="quick-chip" onclick="quickFillExerciseField('modal-ex-weight', 'BW')">BW</button>
          <button type="button" class="quick-chip" onclick="quickFillExerciseField('modal-ex-weight', '10kg')">10kg</button>
          <button type="button" class="quick-chip" onclick="quickFillExerciseField('modal-ex-weight', '15kg')">15kg</button>
          <button type="button" class="quick-chip" onclick="quickFillExerciseField('modal-ex-weight', '20kg')">20kg</button>
          <button type="button" class="quick-chip" onclick="quickFillExerciseField('modal-ex-weight', '30kg')">30kg</button>
        </div>
        <input type="text" id="modal-ex-weight" class="form-input" inputmode="decimal" placeholder="14kg or Bodyweight">
      </div>
      <div class="form-group">
        <label class="form-label">Rest Timer</label>
        <div class="quick-chips-bar">
          <button type="button" class="quick-chip" onclick="quickFillExerciseField('modal-ex-rest', '30s')">30s</button>
          <button type="button" class="quick-chip" onclick="quickFillExerciseField('modal-ex-rest', '60s')">60s</button>
          <button type="button" class="quick-chip" onclick="quickFillExerciseField('modal-ex-rest', '90s')">90s</button>
          <button type="button" class="quick-chip" onclick="quickFillExerciseField('modal-ex-rest', '120s')">120s</button>
          <button type="button" class="quick-chip" onclick="quickFillExerciseField('modal-ex-rest', '180s')">180s</button>
        </div>
        <input type="text" id="modal-ex-rest" class="form-input" inputmode="numeric" placeholder="60s">
      </div>
    </div>

    <div class="form-group">
      <label class="form-label">Coaching Technique Cue</label>
      <textarea id="modal-ex-cue" class="form-textarea" style="min-height:60px;" placeholder="e.g. Keep chest tall, drive through heels, control tempo..."></textarea>
    </div>

    <div class="form-group">
      <label class="form-label">Video URLs (One per line)</label>
      <textarea id="modal-ex-videos" class="form-textarea" style="min-height:55px;" placeholder="https://youtube.com/watch?v=..."></textarea>
    </div>

    <div class="modal-footer-bar">
      <button class="btn-header" onclick="closeModal('exercise-modal')">Cancel</button>
      <button class="btn-header btn-header-primary" onclick="saveExerciseModal()">Done</button>
    </div>
  </div>
</div>

<!-- MODAL: MOBILE PHONE SIMULATOR -->
<div class="modal-overlay" id="preview-modal">
  <div class="modal-window preview-modal-window">
    <div class="preview-modal-controls">
      <div style="display:flex; gap:8px;">
        <button class="btn-header btn-header-primary" id="btn-device-std" onclick="setPreviewDevice('standard')">iPhone</button>
        <button class="btn-header" id="btn-device-max" onclick="setPreviewDevice('wide')">Max</button>
      </div>
      <button class="modal-close-btn" onclick="closeModal('preview-modal')"><svg class="mui-icon" viewBox="0 0 24 24" style="width:18px;height:18px;"><path d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z"/></svg></button>
    </div>
    
    <div class="preview-stage">
      <div class="phone-chassis" id="phone-chassis">
        <div class="phone-notch"></div>
        <iframe id="preview-iframe" class="client-iframe"></iframe>
      </div>
    </div>
  </div>
</div>


<!-- MODAL: BACKUP & RESTORE DATA -->
<div class="modal-overlay" id="backup-modal">
  <div class="modal-window" style="max-width:520px;">
    <div class="modal-header-bar">
      <div class="modal-title">{MUI_ICONS['backup']} Coach Studio Data & Sync</div>
      <button class="modal-close-btn" onclick="closeModal('backup-modal')"><svg class="mui-icon" viewBox="0 0 24 24" style="width:18px;height:18px;"><path d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z"/></svg></button>
    </div>

    <div style="margin:16px 0; display:flex; flex-direction:column; gap:16px;">
      <div style="background:var(--bg-card); border:1px solid var(--border); border-radius:12px; padding:16px;">
        <div style="font-weight:800; color:var(--text-main); margin-bottom:4px; font-size:1.05em;">Backup Studio (.json)</div>
        <div style="font-size:0.82em; color:var(--text-muted); margin-bottom:12px; line-height:1.4;">
          Export a complete snapshot of all athletes, comments, customized routines, master session templates, and custom library exercises.
        </div>
        <button class="btn-header btn-header-primary" onclick="exportStudioBackupJson()" style="width:100%; justify-content:center;">
          {MUI_ICONS['save']} Download Full Backup (.json)
        </button>
      </div>

      <div style="background:var(--bg-card); border:1px solid var(--border); border-radius:12px; padding:16px;">
        <div style="font-weight:800; color:var(--text-main); margin-bottom:4px; font-size:1.05em;">Restore from Backup</div>
        <div style="font-size:0.82em; color:var(--text-muted); margin-bottom:12px; line-height:1.4;">
          Restore athletes, templates, and library from a previously exported backup JSON file.
        </div>
        <input type="file" id="restore-file-input" accept=".json" style="display:none;" onchange="handleRestoreFileSelected(event)">
        <button class="btn-header" onclick="document.getElementById('restore-file-input').click()" style="width:100%; justify-content:center;">
          {MUI_ICONS['folder']} Select Backup File (.json)...
        </button>
      </div>
    </div>

    <div class="modal-footer-bar">
      <button class="btn-header" onclick="closeModal('backup-modal')">Close</button>
    </div>
  </div>
</div>

<!-- MODAL: PASTE FROM EXCEL / CSV -->
<div class="modal-overlay" id="batch-import-modal">
  <div class="modal-window" style="max-width:640px;">
    <div class="modal-header-bar">
      <div class="modal-title">{MUI_ICONS['copy']} Paste from Excel / Google Sheets</div>
      <div style="display:flex; align-items:center; gap:8px;">
        <button type="button" class="btn-header" onclick="downloadCsvTemplate()" style="font-size:0.78em; padding:6px 12px;">
          {MUI_ICONS['download']} Download CSV Template
        </button>
        <button class="modal-close-btn" onclick="closeModal('batch-import-modal')"><svg class="mui-icon" viewBox="0 0 24 24" style="width:18px;height:18px;"><path d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z"/></svg></button>
      </div>
    </div>

    <div style="margin:16px 0; display:flex; flex-direction:column; gap:12px;">
      <div style="font-size:0.82em; color:var(--text-muted); line-height:1.4;">
        Copy rows directly from Excel or Google Sheets and paste them below.<br>
        Supported format (Tab or Comma separated): <strong>Exercise Name [tab] Sets x Reps [tab] Weight [tab] Rest [tab] Coaching Cue [tab] Superset (optional)</strong>
      </div>
      <textarea id="batch-import-text" class="form-textarea" style="min-height:150px; font-family:monospace; font-size:0.85em;" placeholder="Dumbbell Bench Press&#9;4 x 8-10&#9;16kg&#9;90s&#9;Pin shoulder blades&#9;A1&#10;Lat Pulldown&#9;3 x 10-12&#9;45kg&#9;75s&#9;Drive elbows down&#9;A2" oninput="previewBatchImport()"></textarea>
      
      <div id="batch-import-preview-count" style="font-size:0.82em; font-weight:700; color:var(--teal);">0 exercises detected</div>
      <div id="batch-import-preview-list" style="max-height:140px; overflow-y:auto; background:rgba(0,0,0,0.3); border:1px solid var(--border); border-radius:8px; padding:8px 12px; font-size:0.8em; color:var(--text-dim);">
        Paste text above to see preview.
      </div>
    </div>

    <div class="modal-footer-bar">
      <button class="btn-header" onclick="closeModal('batch-import-modal')">Cancel</button>
      <button class="btn-header btn-header-primary" id="btn-confirm-batch-import" onclick="confirmBatchImport()">Insert Exercises</button>
    </div>
  </div>
</div>

<!-- MODAL: EXERCISE LIBRARY -->
<div class="modal-overlay" id="library-modal">
  <div class="modal-window lib-modal-window">
    <div class="modal-header-bar">
      <div style="display:flex; align-items:center; gap:12px;">
        <div class="modal-title">{MUI_ICONS['library']} Exercise Library</div>
        <button class="btn-header btn-header-primary" onclick="openCreateCustomExerciseModal()" style="font-size:0.78em; padding:6px 14px; border-radius:10px;">
          {MUI_ICONS['add']} New Custom Exercise
        </button>
      </div>
      <button class="modal-close-btn" onclick="closeModal('library-modal')"><svg class="mui-icon" viewBox="0 0 24 24" style="width:18px;height:18px;"><path d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z"/></svg></button>
    </div>

    <input type="text" id="lib-search" class="lib-search-input" placeholder="Search library exercises..." oninput="renderLibraryCards()">
    
    <div class="lib-pills-row" id="lib-pills">
      <!-- Filter pills injected dynamically -->
    </div>

    <div class="lib-cards-grid" id="lib-cards">
      <!-- Exercise library cards injected dynamically -->
    </div>
  </div>
</div>

<!-- MODAL: PICK SESSION (FOR CURRENT ATHLETE) -->
<div class="modal-overlay" id="pick-session-modal">
  <div class="modal-window" style="max-width:460px;">
    <div class="modal-header-bar">
      <div>
        <div class="modal-title" id="pick-session-title">Add to Session</div>
        <div style="font-size:0.8em; color:var(--teal); margin-top:2px;" id="pick-session-subtitle">Select target session:</div>
      </div>
      <button class="modal-close-btn" onclick="closeModal('pick-session-modal')"><svg class="mui-icon" viewBox="0 0 24 24" style="width:18px;height:18px;"><path d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z"/></svg></button>
    </div>
    <div id="pick-session-list" style="display:flex; flex-direction:column; gap:8px; max-height:50vh; overflow-y:auto; margin:16px 0;">
      <!-- Session options injected dynamically -->
    </div>
    <div class="modal-footer-bar">
      <button class="btn-header" onclick="closeModal('pick-session-modal')">Cancel</button>
    </div>
  </div>
</div>

<!-- MODAL: GLOBAL ATHLETE & SESSION PICKER -->
<div class="modal-overlay" id="pick-global-modal">
  <div class="modal-window" style="max-width:460px;">
    <div class="modal-header-bar">
      <div>
        <div class="modal-title" id="pick-global-title">Add Exercise to Athlete</div>
        <div style="font-size:0.8em; color:var(--teal); margin-top:2px;" id="pick-global-subtitle">Choose athlete and session</div>
      </div>
      <button class="modal-close-btn" onclick="closeModal('pick-global-modal')"><svg class="mui-icon" viewBox="0 0 24 24" style="width:18px;height:18px;"><path d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z"/></svg></button>
    </div>
    <div style="margin:16px 0; display:flex; flex-direction:column; gap:12px;">
      <div class="form-group">
        <label class="form-label">Select Athlete</label>
        <select id="pick-global-ath-select" class="form-input" onchange="updateGlobalSessionPickerOptions()"></select>
      </div>
      <div class="form-group">
        <label class="form-label">Select Session</label>
        <select id="pick-global-sess-select" class="form-input"></select>
      </div>
    </div>
    <div class="modal-footer-bar">
      <button class="btn-header" onclick="closeModal('pick-global-modal')">Cancel</button>
      <button class="btn-header btn-header-primary" onclick="confirmAddGlobalExercise()">Add Exercise</button>
    </div>
  </div>
</div>

<!-- MODAL: CREATE CUSTOM EXERCISE -->
<div class="modal-overlay" id="create-custom-exercise-modal">
  <div class="modal-window" style="max-width:540px;">
    <div class="modal-header-bar">
      <div class="modal-title">Create Custom Exercise</div>
      <button class="modal-close-btn" onclick="closeModal('create-custom-exercise-modal')"><svg class="mui-icon" viewBox="0 0 24 24" style="width:18px;height:18px;"><path d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z"/></svg></button>
    </div>
    <div class="form-row-2" style="margin-top:10px;">
      <div class="form-group">
        <label class="form-label">Exercise Name *</label>
        <input type="text" id="custom-ex-name" class="form-input" placeholder="e.g. Bulgarian Split Squat">
      </div>
      <div class="form-group">
        <label class="form-label">Category</label>
        <select id="custom-ex-category" class="form-input">
          <option value="Primary Strength">Primary Strength</option>
          <option value="Secondary Strength">Secondary Strength</option>
          <option value="Warm-Up">Warm-Up</option>
          <option value="Accessory">Accessory</option>
          <option value="Core & Stability">Core & Stability</option>
          <option value="Conditioning / Finisher">Conditioning / Finisher</option>
          <option value="Mobility / Cooldown">Mobility / Cooldown</option>
        </select>
      </div>
    </div>
    <div class="form-row-3">
      <div class="form-group">
        <label class="form-label">Target Group</label>
        <select id="custom-ex-group" class="form-input">
          <option value="Upper Body">Upper Body</option>
          <option value="Lower Body">Lower Body</option>
          <option value="Core">Core</option>
          <option value="Warm-Up">Warm-Up</option>
        </select>
      </div>
      <div class="form-group">
        <label class="form-label">Default Sets & Reps</label>
        <input type="text" id="custom-ex-details" class="form-input" placeholder="3 x 8-10">
      </div>
      <div class="form-group">
        <label class="form-label">Default Weight</label>
        <input type="text" id="custom-ex-weight" class="form-input" placeholder="16kg">
      </div>
    </div>
    <div class="form-row-2">
      <div class="form-group">
        <label class="form-label">Rest Timer</label>
        <input type="text" id="custom-ex-rest" class="form-input" placeholder="60s">
      </div>
      <div class="form-group">
        <label class="form-label">Demo Video URL</label>
        <input type="text" id="custom-ex-video" class="form-input" placeholder="https://youtube.com/...">
      </div>
    </div>
    <div class="form-group">
      <label class="form-label">Coach Technique Cue</label>
      <textarea id="custom-ex-cue" class="form-textarea" style="min-height:55px;" placeholder="Key coaching cues, form instructions, breathing tempo..."></textarea>
    </div>
    <div class="modal-footer-bar">
      <button class="btn-header" onclick="closeModal('create-custom-exercise-modal')">Cancel</button>
      <button class="btn-header btn-header-primary" onclick="saveCustomExercise()">{MUI_ICONS['save']} Save to Library</button>
    </div>
  </div>
</div>

<!-- MODAL: ASSIGN TEMPLATE TO ATHLETES -->
<div class="modal-overlay" id="assign-template-modal">
  <div class="modal-window" style="max-width:500px;">
    <div class="modal-header-bar">
      <div>
        <div class="modal-title" id="assign-template-title">{MUI_ICONS['rocket']} Assign Template to Athletes</div>
        <div style="font-size:0.8em; color:var(--teal); margin-top:2px;" id="assign-template-subtitle">Choose session name and athletes</div>
      </div>
      <button class="modal-close-btn" onclick="closeModal('assign-template-modal')"><svg class="mui-icon" viewBox="0 0 24 24" style="width:18px;height:18px;"><path d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z"/></svg></button>
    </div>

    <div style="margin:16px 0; display:flex; flex-direction:column; gap:14px;">
      <div class="form-group">
        <label class="form-label">Session Name for Athletes *</label>
        <input type="text" id="assign-custom-session-name" class="form-input" placeholder="e.g. Session 3 - Push Power">
        <span style="font-size:0.75em; color:var(--text-dim); margin-top:4px;">You can name this workout whatever you want for each athlete's program.</span>
      </div>

      <div class="form-group">
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:8px;">
          <label class="form-label" style="margin-bottom:0;">Select Athletes *</label>
          <button type="button" class="btn-header" onclick="toggleSelectAllAssignAthletes()" style="font-size:0.75em; padding:4px 10px;">Select All</button>
        </div>
        <div id="assign-athletes-list" style="display:flex; flex-direction:column; gap:6px; max-height:220px; overflow-y:auto; padding-right:4px;">
          <!-- Athletes with checkboxes injected dynamically -->
        </div>
      </div>
    </div>

    <div class="modal-footer-bar">
      <button class="btn-header" onclick="closeModal('assign-template-modal')">Cancel</button>
      <button class="btn-header btn-header-primary" id="btn-confirm-assign" onclick="confirmAssignTemplateToAthletes()">{MUI_ICONS['rocket']} Assign Session</button>
    </div>
  </div>
</div>

<!-- MODAL: CREATE MASTER TEMPLATE -->
<div class="modal-overlay" id="create-template-modal">
  <div class="modal-window" style="max-width:480px;">
    <div class="modal-header-bar">
      <div class="modal-title">{MUI_ICONS['add']} Create Session Template</div>
      <button class="modal-close-btn" onclick="closeModal('create-template-modal')"><svg class="mui-icon" viewBox="0 0 24 24" style="width:18px;height:18px;"><path d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z"/></svg></button>
    </div>
    <div style="margin:16px 0; display:flex; flex-direction:column; gap:12px;">
      <div class="form-group">
        <label class="form-label">Template Name *</label>
        <input type="text" id="template-name-input" class="form-input" placeholder="e.g. Upper Body Hypertrophy">
      </div>
      <div class="form-group">
        <label class="form-label">Category / Focus</label>
        <select id="template-category-input" class="form-input">
          <option value="Push Focus">Push (Chest, Delts, Triceps)</option>
          <option value="Pull Focus">Pull (Back, Biceps)</option>
          <option value="Leg Focus">Lower Body (Quads, Posterior)</option>
          <option value="Full Body">Full Body</option>
          <option value="Conditioning & Core">Conditioning & Core</option>
          <option value="Mobility & Recovery">Mobility & Recovery</option>
        </select>
      </div>
    </div>
    <div class="modal-footer-bar">
      <button class="btn-header" onclick="closeModal('create-template-modal')">Cancel</button>
      <button class="btn-header btn-header-primary" onclick="confirmCreateTemplate()">Create Template</button>
    </div>
  </div>
</div>

<!-- MODAL: CALENDAR DAY LOG & WORKOUT INSPECTOR -->
<div class="modal-overlay" id="calendar-day-modal">
  <div class="modal-window cal-day-modal-window">
    <div class="modal-header-bar">
      <div style="display:flex; align-items:center; gap:10px;">
        <div class="modal-title">{MUI_ICONS['calendar']} <span id="cal-day-modal-title">Workout Details</span></div>
        <span class="athlete-badge-count" id="cal-day-modal-count-badge">0 Workouts</span>
      </div>
      <button class="modal-close-btn" onclick="closeModal('calendar-day-modal')"><svg class="mui-icon" viewBox="0 0 24 24" style="width:18px;height:18px;"><path d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z"/></svg></button>
    </div>

    <div style="margin:16px 0; overflow-y:auto; padding-right:4px; display:flex; flex-direction:column; gap:14px; flex:1;">
      <!-- SECTION 1: Assigned & Logged Workouts on this Day -->
      <div id="cal-day-workouts-container" style="display:flex; flex-direction:column; gap:10px;">
        <!-- Injected dynamically with full exercise breakdown & Builder edit buttons -->
      </div>

      <!-- SECTION 2: Collapsible Assign / Schedule New Workout Box -->
      <div style="background:rgba(0,0,0,0.3); border:1px solid var(--border); border-radius:12px; padding:14px;">
        <div style="display:flex; justify-content:space-between; align-items:center; cursor:pointer;" onclick="toggleCalAssignBox()">
          <div style="font-size:0.85em; font-weight:800; color:var(--text-main); display:flex; align-items:center; gap:8px;">
            {MUI_ICONS['add']} <span id="cal-assign-box-title">Assign or Log a Workout to this Date</span>
          </div>
          <button type="button" class="btn-header" style="font-size:0.75em; padding:3px 8px;" id="btn-toggle-cal-assign">
            {MUI_ICONS['add']} Expand
          </button>
        </div>

        <div id="cal-assign-box-body" style="display:none; margin-top:14px; flex-direction:column; gap:12px;">
          <div class="form-group">
            <label class="form-label">Date</label>
            <input type="date" id="cal-log-date" class="form-input">
          </div>

          <div class="cal-assign-grid">
            <div class="form-group">
              <label class="form-label">Athlete</label>
              <select id="cal-log-athlete" class="form-input" onchange="onCalLogAthleteChange()">
                <!-- Populated dynamically -->
              </select>
            </div>

            <div class="form-group">
              <label class="form-label">Workout Routine / Template</label>
              <select id="cal-log-session" class="form-input">
                <!-- Populated dynamically -->
              </select>
            </div>
          </div>

          <div class="form-group">
            <label class="form-label">Status</label>
            <div class="status-pill-group">
              <button type="button" class="status-pill-btn active completed" onclick="setCalLogStatus('completed', this)">Completed</button>
              <button type="button" class="status-pill-btn scheduled" onclick="setCalLogStatus('scheduled', this)">Scheduled</button>
              <button type="button" class="status-pill-btn missed" onclick="setCalLogStatus('missed', this)">Missed</button>
            </div>
            <input type="hidden" id="cal-log-status-val" value="completed">
          </div>

          <div class="form-group">
            <label class="form-label">Workout Notes / Instructions</label>
            <textarea id="cal-log-notes" class="form-textarea" style="min-height:70px;" placeholder="e.g. Focus on explosive tempo during bench press."></textarea>
          </div>

          <div style="display:flex; justify-content:flex-end;">
            <button class="btn-header btn-header-primary cal-assign-submit-btn" onclick="saveCalendarDayLog()">{MUI_ICONS['save']} Assign Workout to Date</button>
          </div>
        </div>
      </div>
    </div>

    <div class="modal-footer-bar">
      <button class="btn-header" onclick="closeModal('calendar-day-modal')">Close</button>
    </div>
  </div>
</div>

<!-- MODAL: IMPORT WORKOUT RECAP -->
<div class="modal-overlay" id="import-recap-modal">
  <div class="modal-window" style="max-width:580px;">
    <div class="modal-header-bar">
      <div class="modal-title">{MUI_ICONS['copy']} Import Workout Recap</div>
      <button class="modal-close-btn" onclick="closeModal('import-recap-modal')"><svg class="mui-icon" viewBox="0 0 24 24" style="width:18px;height:18px;"><path d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z"/></svg></button>
    </div>

    <div style="margin:16px 0; display:flex; flex-direction:column; gap:12px;">
      <div style="font-size:0.82em; color:var(--text-muted); line-height:1.4;">
        Paste the workout recap copied from the athlete's client app (or received via WhatsApp) below:
      </div>

      <textarea id="import-recap-text" class="form-textarea" style="min-height:120px; font-family:monospace; font-size:0.84em;" placeholder="Workout Recap: Jada Said&#10;Session 1 - Upper Power (2026-09-06)&#10;------------------------&#10;Bent-Elbow Band Pull-Aparts: 2 sets (Light Band x 10, Light Band x 10)&#10;DB Bench Press: 4 sets (14kg x 10, 14kg x 10, 14kg x 10, 14kg x 8)&#10;&#10;Coach Notes: Felt strong on rows today." oninput="previewImportedRecap()"></textarea>

      <!-- Live Parse Preview Card -->
      <div id="import-recap-preview-wrap" style="display:none; background:rgba(0,0,0,0.35); border:1px solid var(--border); border-radius:12px; padding:14px;">
        <div style="font-size:0.8em; font-weight:800; color:var(--teal); text-transform:uppercase; letter-spacing:0.04em; margin-bottom:10px; display:flex; align-items:center; gap:6px;">
          {MUI_ICONS['check']} Parsed Workout Summary
        </div>

        <div style="display:grid; grid-template-columns:1fr 1fr; gap:10px; margin-bottom:10px;">
          <div>
            <label class="form-label" style="font-size:0.75em;">Matched Athlete</label>
            <select id="import-recap-ath-select" class="form-input" style="font-size:0.85em; padding:6px 10px;">
              <!-- Populated dynamically -->
            </select>
          </div>
          <div>
            <label class="form-label" style="font-size:0.75em;">Workout Date</label>
            <input type="date" id="import-recap-date-input" class="form-input" style="font-size:0.85em; padding:6px 10px;">
          </div>
        </div>

        <div style="margin-bottom:8px;">
          <div style="font-size:0.75em; color:var(--text-muted); font-weight:600;">Session:</div>
          <div id="import-recap-session-title" style="font-size:0.9em; font-weight:800; color:var(--text-main); margin-top:2px;"></div>
        </div>

        <div style="margin-bottom:8px;">
          <div style="font-size:0.75em; color:var(--text-muted); font-weight:600;">Exercises Detected:</div>
          <div id="import-recap-exercises-list" style="margin-top:4px; max-height:90px; overflow-y:auto; display:flex; flex-direction:column; gap:4px;"></div>
        </div>

        <div id="import-recap-notes-wrap" style="margin-top:6px;">
          <div style="font-size:0.75em; color:var(--text-muted); font-weight:600;">Athlete Notes:</div>
          <div id="import-recap-notes-text" style="font-size:0.82em; color:var(--text-dim); margin-top:2px; font-style:italic;"></div>
        </div>
      </div>
    </div>

    <div class="modal-footer-bar">
      <button class="btn-header" onclick="closeModal('import-recap-modal')">Cancel</button>
      <button class="btn-header btn-header-primary" id="btn-confirm-import-recap" onclick="confirmImportRecap()">{MUI_ICONS['save']} Import to Calendar</button>
    </div>
  </div>
</div>

<!-- MODAL: MOBILE TOOLS BOTTOM SHEET -->
<div class="modal-overlay" id="mobile-tools-modal">
  <div class="modal-window" style="max-width:460px;">
    <div class="sheet-drag-pill"></div>
    <div class="modal-header-bar">
      <div class="modal-title">{MUI_ICONS['settings']} Studio Tools & Settings</div>
      <button class="modal-close-btn" onclick="closeModal('mobile-tools-modal')">
        <svg class="mui-icon" viewBox="0 0 24 24" style="width:18px;height:18px;"><path d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z"/></svg>
      </button>
    </div>

    <div style="display:flex; flex-direction:column; gap:8px; margin:14px 0;">
      <button type="button" class="btn-header" onclick="closeModal('mobile-tools-modal'); openCloudSyncModal();" style="width:100%; justify-content:flex-start; padding:12px 14px; font-size:0.9em; gap:12px;">
        <span style="color:var(--teal);">{MUI_ICONS['cloud_sync']}</span>
        <span>24/7 Cloud Sync (GitHub Gist)</span>
      </button>
      <button type="button" class="btn-header" onclick="closeModal('mobile-tools-modal'); lockStudioNow();" style="width:100%; justify-content:flex-start; padding:12px 14px; font-size:0.9em; gap:12px;">
        <span style="color:var(--teal);">{MUI_ICONS['lock']}</span>
        <span>Lock Studio Now</span>
      </button>
      <button type="button" class="btn-header" onclick="closeModal('mobile-tools-modal'); openPasscodeSettingsModal();" style="width:100%; justify-content:flex-start; padding:12px 14px; font-size:0.9em; gap:12px;">
        <span style="color:var(--teal);">{MUI_ICONS['shield']}</span>
        <span>Passcode Security Settings</span>
      </button>
      <button type="button" class="btn-header" onclick="closeModal('mobile-tools-modal'); openNewAthleteModal();" style="width:100%; justify-content:flex-start; padding:12px 14px; font-size:0.9em; gap:12px;">
        <span style="color:var(--teal);">{MUI_ICONS['add']}</span>
        <span>Create New Athlete</span>
      </button>
      <button type="button" class="btn-header" onclick="closeModal('mobile-tools-modal'); openImportRecapModal();" style="width:100%; justify-content:flex-start; padding:12px 14px; font-size:0.9em; gap:12px;">
        <span style="color:var(--teal);">{MUI_ICONS['copy']}</span>
        <span>Import Workout Recap</span>
      </button>
      <button type="button" class="btn-header" onclick="closeModal('mobile-tools-modal'); openLibraryModal();" style="width:100%; justify-content:flex-start; padding:12px 14px; font-size:0.9em; gap:12px;">
        <span style="color:var(--teal);">{MUI_ICONS['library']}</span>
        <span>Exercise Library</span>
      </button>
      <button type="button" class="btn-header" onclick="closeModal('mobile-tools-modal'); openBackupModal();" style="width:100%; justify-content:flex-start; padding:12px 14px; font-size:0.9em; gap:12px;">
        <span style="color:var(--teal);">{MUI_ICONS['backup']}</span>
        <span>Backup & Sync (.json)</span>
      </button>
      <button type="button" class="btn-header" onclick="closeModal('mobile-tools-modal'); downloadCsvTemplate();" style="width:100%; justify-content:flex-start; padding:12px 14px; font-size:0.9em; gap:12px;">
        <span style="color:var(--teal);">{MUI_ICONS['download']}</span>
        <span>Download CSV Template</span>
      </button>
    </div>

    <div class="modal-footer-bar">
      <button class="btn-header" onclick="closeModal('mobile-tools-modal')" style="width:100%; justify-content:center;">Close</button>
    </div>
  </div>
</div>

<!-- NATIVE MOBILE BOTTOM NAVIGATION BAR -->
<nav class="mobile-bottom-nav" id="mobile-bottom-nav">
  <button type="button" class="mobile-nav-item active" id="mobile-nav-athletes" onclick="navigateTo('hub')">
    <span class="mobile-nav-icon">{MUI_ICONS['athletes']}</span>
    <span class="mobile-nav-label">Athletes</span>
  </button>
  <button type="button" class="mobile-nav-item" id="mobile-nav-calendar" onclick="navigateTo('calendar')">
    <span class="mobile-nav-icon">{MUI_ICONS['calendar']}</span>
    <span class="mobile-nav-label">Calendar</span>
  </button>
  <button type="button" class="mobile-nav-item" id="mobile-nav-templates" onclick="navigateTo('templates')">
    <span class="mobile-nav-icon">{MUI_ICONS['templates']}</span>
    <span class="mobile-nav-label">Templates</span>
  </button>
  <button type="button" class="mobile-nav-item" id="mobile-nav-tools" onclick="openMobileToolsSheet()">
    <span class="mobile-nav-icon">{MUI_ICONS['settings']}</span>
    <span class="mobile-nav-label">Tools</span>
  </button>
</nav>

"""

STUDIO_JS = """

// -------------------------------------------------------------
// UNIVERSAL CUSTOM APP-STYLED SELECT COMPONENT
// -------------------------------------------------------------
function enhanceSelect(selectEl) {
  if (!selectEl || selectEl.dataset.customized === 'true') {
    if (selectEl && selectEl._syncCustomSelect) selectEl._syncCustomSelect();
    return;
  }
  selectEl.dataset.customized = 'true';
  selectEl.style.display = 'none';

  const wrapper = document.createElement('div');
  wrapper.className = 'custom-select-wrapper';
  if (selectEl.style.width) wrapper.style.width = selectEl.style.width;
  if (selectEl.classList.contains('form-input')) wrapper.style.width = selectEl.style.width || '100%';

  const trigger = document.createElement('button');
  trigger.type = 'button';
  trigger.className = 'custom-select-trigger';

  const triggerText = document.createElement('span');
  triggerText.className = 'custom-select-trigger-text';

  const arrow = document.createElement('span');
  arrow.className = 'custom-select-arrow';
  arrow.innerHTML = `<svg class="mui-icon" viewBox="0 0 24 24" style="width:18px;height:18px;"><path d="M7 10l5 5 5-5z"/></svg>`;

  trigger.appendChild(triggerText);
  trigger.appendChild(arrow);

  const menu = document.createElement('div');
  menu.className = 'custom-select-menu';

  function updateOptions() {
    menu.innerHTML = '';
    const selectedOpt = selectEl.options[selectEl.selectedIndex] || selectEl.options[0];
    triggerText.textContent = selectedOpt ? selectedOpt.textContent : '';

    Array.from(selectEl.options).forEach((opt, idx) => {
      const item = document.createElement('div');
      item.className = 'custom-select-item' + (opt.selected ? ' selected' : '');
      item.innerHTML = `<span>${escapeHtml(opt.textContent)}</span>${opt.selected ? MUI.check : ''}`;
      item.onclick = function(e) {
        e.stopPropagation();
        selectEl.selectedIndex = idx;
        selectEl.value = opt.value;
        selectEl.dispatchEvent(new Event('change', { bubbles: true }));
        updateOptions();
        closeAllCustomSelects();
      };
      menu.appendChild(item);
    });
  }

  trigger.onclick = function(e) {
    e.stopPropagation();
    const isOpen = menu.classList.contains('open');
    closeAllCustomSelects();
    if (!isOpen) {
      updateOptions();
      menu.classList.add('open');
      trigger.classList.add('active');
    }
  };

  selectEl.parentNode.insertBefore(wrapper, selectEl);
  wrapper.appendChild(selectEl);
  wrapper.appendChild(trigger);
  wrapper.appendChild(menu);

  selectEl.addEventListener('change', updateOptions);
  updateOptions();

  selectEl._syncCustomSelect = updateOptions;
}

function closeAllCustomSelects() {
  document.querySelectorAll('.custom-select-menu.open').forEach(m => m.classList.remove('open'));
  document.querySelectorAll('.custom-select-trigger.active').forEach(t => t.classList.remove('active'));
}

document.addEventListener('click', closeAllCustomSelects);

function initAllCustomSelects() {
  document.querySelectorAll('select.form-input').forEach(enhanceSelect);
}


// -------------------------------------------------------------
// HEADER SWITCHER & STUDIO TOOLS MENU
// -------------------------------------------------------------
function toggleStudioToolsMenu(e) {
  e.stopPropagation();
  const menu = document.getElementById('studio-tools-menu');
  const btn = document.getElementById('btn-studio-tools-menu');
  if (!menu) return;
  const isOpen = menu.classList.contains('open');
  closeAllCustomSelects();
  if (!isOpen) {
    menu.classList.add('open');
    if (btn) btn.classList.add('active');
  }
}

function toggleAthleteSwitcher(e) {
  e.stopPropagation();
  const menu = document.getElementById('header-athlete-menu');
  const btn = document.getElementById('header-athlete-btn');
  if (!menu) return;
  const isOpen = menu.classList.contains('open');
  closeAllCustomSelects();
  if (!isOpen) {
    menu.classList.add('open');
    if (btn) btn.classList.add('active');
  }
}

function toggleSessionSwitcher(e) {
  e.stopPropagation();
  const menu = document.getElementById('header-session-menu');
  const btn = document.getElementById('header-session-btn');
  if (!menu) return;
  const isOpen = menu.classList.contains('open');
  closeAllCustomSelects();
  if (!isOpen) {
    menu.classList.add('open');
    if (btn) btn.classList.add('active');
  }
}

// State
let studioAthletes = [];
try {
  const saved = localStorage.getItem('coach_studio_athletes');
  if (saved) {
    studioAthletes = JSON.parse(saved);
  }
} catch (e) {}

if (!studioAthletes || !Array.isArray(studioAthletes) || studioAthletes.length === 0 || !studioAthletes[0].name) {
  studioAthletes = window.INITIAL_ATHLETES || [];
  try { localStorage.setItem('coach_studio_athletes', JSON.stringify(studioAthletes)); } catch (e) {}
}

// Ensure every athlete has a logs array
studioAthletes.forEach(a => {
  if (!a.logs || !Array.isArray(a.logs)) a.logs = [];
});

let calendarCurrentDate = new Date(2026, 8, 6); // September 2026
let calendarSelectedAthleteId = 'all';
let parsedRecapCache = null;

let activeAthleteId = studioAthletes[0] ? studioAthletes[0].id : null;
let currentView = 'hub'; // 'hub' | 'athlete' | 'builder'
let activeSessionIndex = 0;
let modalEditingCardIdx = null;
let sessionToCopyIdx = null;
let confirmCallback = null;
let promptCallback = null;

// Built-in library exercises
const DEFAULT_LIBRARY = [
  { group: 'Warm-Up', category: 'Warm-Up', exercise: "World's Greatest Stretch", details: '2 x 5 per side', weight: 'Bodyweight', rest: '30s', rest_seconds: 30, total_sets: 2, cue: 'Open thoracic spine tall toward ceiling.', videos: [] },
  { group: 'Warm-Up', category: 'Warm-Up', exercise: 'Cat-Cow Mobility', details: '2 x 10', weight: 'Bodyweight', rest: '30s', rest_seconds: 30, total_sets: 2, cue: 'Coordinate movement smoothly with breathing.', videos: [] },
  { group: 'Warm-Up', category: 'Warm-Up', exercise: 'Glute Bridges', details: '2 x 12', weight: 'Bodyweight', rest: '30s', rest_seconds: 30, total_sets: 2, cue: 'Squeeze glutes at top without arching lower back.', videos: [] },
  { group: 'Upper Body', category: 'Primary Strength', exercise: 'Dumbbell Bench Press', details: '4 x 8-10', weight: '16kg', rest: '90s', rest_seconds: 90, total_sets: 4, cue: 'Drive feet into floor and keep shoulder blades pinned.', videos: [] },
  { group: 'Upper Body', category: 'Primary Strength', exercise: 'Dumbbell Incline Bench Press', details: '3 x 10-12', weight: '14kg', rest: '90s', rest_seconds: 90, total_sets: 3, cue: 'Control the descent and press on slight angle.', videos: [] },
  { group: 'Upper Body', category: 'Primary Strength', exercise: 'Barbell Overhead Press', details: '4 x 6-8', weight: '40kg', rest: '90s', rest_seconds: 90, total_sets: 4, cue: 'Lock glutes and core tight at lockout.', videos: [] },
  { group: 'Upper Body', category: 'Secondary Strength', exercise: 'DB Bent-Over Row', details: '4 x 8-10', weight: '18kg', rest: '90s', rest_seconds: 90, total_sets: 4, cue: 'Pull elbow towards hip pocket keeping spine flat.', videos: [] },
  { group: 'Upper Body', category: 'Secondary Strength', exercise: 'Lat Pulldown', details: '3 x 10-12', weight: '45kg', rest: '75s', rest_seconds: 75, total_sets: 3, cue: 'Drive elbows down towards hips.', videos: [] },
  { group: 'Lower Body', category: 'Primary Strength', exercise: 'Barbell Back Squat', details: '4 x 6-8', weight: '80kg', rest: '120s', rest_seconds: 120, total_sets: 4, cue: 'Hit parallel depth with knees tracking over toes.', videos: [] },
  { group: 'Lower Body', category: 'Primary Strength', exercise: 'Romanian Deadlift (RDL)', details: '3 x 10', weight: '60kg', rest: '90s', rest_seconds: 90, total_sets: 3, cue: 'Hinge hips back and feel deep hamstring stretch.', videos: [] },
  { group: 'Lower Body', category: 'Primary Strength', exercise: 'Bulgarian Split Squat', details: '3 x 8 per leg', weight: '12kg DBs', rest: '90s', rest_seconds: 90, total_sets: 3, cue: 'Keep torso slightly forward, drive up through front heel.', videos: [] },
  { group: 'Lower Body', category: 'Secondary Strength', exercise: 'Leg Press', details: '3 x 12-15', weight: '120kg', rest: '90s', rest_seconds: 90, total_sets: 3, cue: 'Controlled tempo, do not let lower back round.', videos: [] },
  { group: 'Core', category: 'Core & Stability', exercise: 'Pallof Press', details: '3 x 12 per side', weight: '10kg', rest: '45s', rest_seconds: 45, total_sets: 3, cue: 'Resist rotational twist, keep core locked.', videos: [] },
  { group: 'Core', category: 'Core & Stability', exercise: 'Hanging Leg Raises', details: '3 x 12', weight: 'Bodyweight', rest: '60s', rest_seconds: 60, total_sets: 3, cue: 'Curl pelvis up towards ribs without swinging.', videos: [] }
];

// MASTER SESSION TEMPLATES DATA
const DEFAULT_TEMPLATES = [
  {
    id: 'tpl_push_1',
    name: 'Push - Chest, Shoulders & Triceps',
    category: 'Push Focus',
    cards: [
      { category: 'Warm-Up', exercise: "World's Greatest Stretch", details: '2 x 5 per side', weight: 'Bodyweight', rest: '30s', rest_seconds: 30, total_sets: 2, cue: 'Open thoracic spine tall toward ceiling.', videos: [] },
      { category: 'Primary Strength', exercise: 'Dumbbell Bench Press', details: '4 x 8-10', weight: '16kg', rest: '90s', rest_seconds: 90, total_sets: 4, cue: 'Drive feet into floor and keep shoulder blades pinned.', videos: [] },
      { category: 'Primary Strength', exercise: 'Dumbbell Incline Bench Press', details: '3 x 10-12', weight: '14kg', rest: '90s', rest_seconds: 90, total_sets: 3, cue: 'Control the descent and press on slight angle.', videos: [] },
      { category: 'Primary Strength', exercise: 'Barbell Overhead Press', details: '4 x 6-8', weight: '40kg', rest: '90s', rest_seconds: 90, total_sets: 4, cue: 'Lock glutes and core tight at lockout.', videos: [] }
    ]
  },
  {
    id: 'tpl_pull_1',
    name: 'Pull - Back, Lats & Posterior Chain',
    category: 'Pull Focus',
    cards: [
      { category: 'Warm-Up', exercise: 'Cat-Cow Mobility', details: '2 x 10', weight: 'Bodyweight', rest: '30s', rest_seconds: 30, total_sets: 2, cue: 'Coordinate movement smoothly with breathing.', videos: [] },
      { category: 'Secondary Strength', exercise: 'DB Bent-Over Row', details: '4 x 8-10', weight: '18kg', rest: '90s', rest_seconds: 90, total_sets: 4, cue: 'Pull elbow towards hip pocket keeping spine flat.', videos: [] },
      { category: 'Secondary Strength', exercise: 'Lat Pulldown', details: '3 x 10-12', weight: '45kg', rest: '75s', rest_seconds: 75, total_sets: 3, cue: 'Drive elbows down towards hips.', videos: [] },
      { category: 'Primary Strength', exercise: 'Romanian Deadlift (RDL)', details: '3 x 10', weight: '60kg', rest: '90s', rest_seconds: 90, total_sets: 3, cue: 'Hinge hips back and feel deep hamstring stretch.', videos: [] }
    ]
  },
  {
    id: 'tpl_legs_1',
    name: 'Lower Body - Quad & Posterior Chain',
    category: 'Leg Focus',
    cards: [
      { category: 'Warm-Up', exercise: 'Glute Bridges', details: '2 x 12', weight: 'Bodyweight', rest: '30s', rest_seconds: 30, total_sets: 2, cue: 'Squeeze glutes at top without arching lower back.', videos: [] },
      { category: 'Primary Strength', exercise: 'Barbell Back Squat', details: '4 x 6-8', weight: '80kg', rest: '120s', rest_seconds: 120, total_sets: 4, cue: 'Hit parallel depth with knees tracking over toes.', videos: [] },
      { category: 'Primary Strength', exercise: 'Bulgarian Split Squat', details: '3 x 8 per leg', weight: '12kg DBs', rest: '90s', rest_seconds: 90, total_sets: 3, cue: 'Keep torso slightly forward, drive up through front heel.', videos: [] },
      { category: 'Secondary Strength', exercise: 'Leg Press', details: '3 x 12-15', weight: '120kg', rest: '90s', rest_seconds: 90, total_sets: 3, cue: 'Controlled tempo, do not let lower back round.', videos: [] }
    ]
  },
  {
    id: 'tpl_fullbody_1',
    name: 'Full Body Conditioning & Core',
    category: 'Full Body',
    cards: [
      { category: 'Warm-Up', exercise: "World's Greatest Stretch", details: '2 x 5 per side', weight: 'Bodyweight', rest: '30s', rest_seconds: 30, total_sets: 2, cue: 'Open thoracic spine tall toward ceiling.', videos: [] },
      { category: 'Primary Strength', exercise: 'Dumbbell Bench Press', details: '3 x 10', weight: '14kg', rest: '60s', rest_seconds: 60, total_sets: 3, cue: 'Controlled tempo.', videos: [] },
      { category: 'Core & Stability', exercise: 'Pallof Press', details: '3 x 12 per side', weight: '10kg', rest: '45s', rest_seconds: 45, total_sets: 3, cue: 'Resist rotational twist, keep core locked.', videos: [] },
      { category: 'Core & Stability', exercise: 'Hanging Leg Raises', details: '3 x 12', weight: 'Bodyweight', rest: '60s', rest_seconds: 60, total_sets: 3, cue: 'Curl pelvis up towards ribs without swinging.', videos: [] }
    ]
  }
];

let sessionTemplates = [];
try {
  const savedTpls = localStorage.getItem('coach_studio_templates');
  if (savedTpls) sessionTemplates = JSON.parse(savedTpls);
} catch (e) {}

if (!sessionTemplates || !Array.isArray(sessionTemplates) || sessionTemplates.length === 0) {
  sessionTemplates = DEFAULT_TEMPLATES;
  try { localStorage.setItem('coach_studio_templates', JSON.stringify(sessionTemplates)); } catch (e) {}
}

function saveTemplatesToStorage() {
  try {
    localStorage.setItem('coach_studio_templates', JSON.stringify(sessionTemplates));
    localStorage.setItem('coach_studio_last_modified', new Date().toISOString());
    if (typeof scheduleCloudAutoPush === 'function') scheduleCloudAutoPush();
  } catch (e) {}
}

function saveAthletesToStorage() {
  try {
    localStorage.setItem('coach_studio_athletes', JSON.stringify(studioAthletes));
    localStorage.setItem('coach_studio_last_modified', new Date().toISOString());
    if (typeof scheduleCloudAutoPush === 'function') scheduleCloudAutoPush();
  } catch (e) {
    console.error('Storage save error:', e);
  }
}

function getActiveAthlete() {
  return studioAthletes.find(a => a.id === activeAthleteId) || studioAthletes[0];
}

function escapeHtml(text) {
  if (!text) return '';
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

// IN-APP NOTIFICATIONS (No native alerts)
function showToast(message, icon = MUI.check) {
  const toast = document.getElementById('toast');
  const msg = document.getElementById('toast-message');
  const ic = document.getElementById('toast-icon');
  if (!toast || !msg) return;
  
  msg.textContent = message;
  if (ic) ic.innerHTML = icon || MUI.check;
  toast.classList.add('show');
  
  clearTimeout(window._toastTimeout);
  window._toastTimeout = setTimeout(() => {
    toast.classList.remove('show');
  }, 2200);
}

function showConfirmModal(title, message, callback, icon = MUI.warning) {
  document.getElementById('confirm-title').textContent = title;
  document.getElementById('confirm-message').textContent = message;
  const ic = document.getElementById('confirm-icon');
  if (ic) ic.innerHTML = icon || MUI.warning;
  confirmCallback = callback;
  openModal('custom-confirm-modal');
}

function closeConfirmModal(confirmed) {
  closeModal('custom-confirm-modal');
  if (confirmCallback) {
    confirmCallback(confirmed);
    confirmCallback = null;
  }
}

function showPromptModal(title, label, defaultValue, callback) {
  document.getElementById('prompt-modal-title').textContent = title;
  document.getElementById('prompt-modal-label').textContent = label;
  const input = document.getElementById('prompt-modal-input');
  input.value = defaultValue || '';
  promptCallback = callback;
  openModal('custom-prompt-modal');
  setTimeout(() => { input.focus(); input.select(); }, 150);
}

function closePromptModal(submitted) {
  const input = document.getElementById('prompt-modal-input');
  const val = input.value.trim();
  closeModal('custom-prompt-modal');
  if (promptCallback) {
    promptCallback(submitted ? val : null);
    promptCallback = null;
  }
}

function openModal(id) {
  const el = document.getElementById(id);
  if (el) el.classList.add('show');
}
function closeModal(id) {
  const el = document.getElementById(id);
  if (el) el.classList.remove('show');
}

function openMobileToolsSheet() {
  openModal('mobile-tools-modal');
}

function quickFillExerciseField(fieldId, val) {
  const el = document.getElementById(fieldId);
  if (el) {
    el.value = val;
    el.focus();
  }
}

// NAVIGATION CONTROLLER
function navigateTo(view, athleteId = null, sessionIdx = 0) {
  if (athleteId) activeAthleteId = athleteId;
  activeSessionIndex = sessionIdx;
  currentView = view;

  document.querySelectorAll('.view-screen').forEach(s => s.classList.remove('active'));
  const activeScreen = document.getElementById('view-' + view);
  if (activeScreen) activeScreen.classList.add('active');

  const tabAthletes = document.getElementById('tab-nav-athletes');
  const tabCalendar = document.getElementById('tab-nav-calendar');
  const tabTemplates = document.getElementById('tab-nav-templates');
  if (tabAthletes) tabAthletes.classList.toggle('active', view === 'hub');
  if (tabCalendar) tabCalendar.classList.toggle('active', view === 'calendar');
  if (tabTemplates) tabTemplates.classList.toggle('active', view === 'templates');

  // Sync Mobile Bottom Navigation active states
  const mobAthletes = document.getElementById('mobile-nav-athletes');
  const mobCalendar = document.getElementById('mobile-nav-calendar');
  const mobTemplates = document.getElementById('mobile-nav-templates');
  const mobTools = document.getElementById('mobile-nav-tools');
  if (mobAthletes) mobAthletes.classList.toggle('active', view === 'hub' || view === 'athlete' || view === 'builder');
  if (mobCalendar) mobCalendar.classList.toggle('active', view === 'calendar');
  if (mobTemplates) mobTemplates.classList.toggle('active', view === 'templates');
  if (mobTools) mobTools.classList.remove('active');

  window.scrollTo({ top: 0, behavior: 'instant' });

  const ath = getActiveAthlete();
  const navBreadcrumb = document.getElementById('nav-breadcrumb');
  const headerActions = document.getElementById('header-actions');

  if (view === 'hub') {
    navBreadcrumb.innerHTML = `<span>/</span> <span style="color:var(--text-main); font-weight:700;">Athletes Directory</span>`;
    headerActions.innerHTML = `
      <button class="btn-header" onclick="navigateTo('calendar')">
        ${MUI.calendar} Calendar
      </button>
      <button class="btn-header" onclick="openLibraryModal()">
        ${MUI.library} Exercise Library
      </button>
    `;
    renderAthletesHub();
  } else if (view === 'calendar') {
    navBreadcrumb.innerHTML = `<span>/</span> <span style="color:var(--text-main); font-weight:700;">Athlete Training Calendar</span>`;
    headerActions.innerHTML = `
      <button class="btn-header" onclick="openImportRecapModal()">
        ${MUI.copy} Import Recap
      </button>
      <button class="btn-header btn-header-primary" onclick="openCalendarDayModal()">
        ${MUI.add} Log Workout
      </button>
    `;
    renderCalendarView();
  } else if (view === 'templates') {
    navBreadcrumb.innerHTML = `<span>/</span> <span style="color:var(--text-main); font-weight:700;">Master Session Templates</span>`;
    headerActions.innerHTML = `
      <button class="btn-header" onclick="openLibraryModal()">
        ${MUI.library} Exercise Library
      </button>
    `;
    renderTemplatesView();
  } else if (view === 'athlete') {
    // Athlete Switcher dropdown
    const athListHtml = studioAthletes.map(a => `
      <div class="custom-select-item ${a.id === ath.id ? 'selected' : ''}" onclick="navigateTo('athlete', '${a.id}'); closeAllCustomSelects();">
        <span>${escapeHtml(a.name)}</span>
        ${a.id === ath.id ? MUI.check : ''}
      </div>
    `).join('');

    navBreadcrumb.innerHTML = `
      <div class="header-switcher-wrap">
        <span class="header-switcher-sep">›</span>
        <button class="switcher-btn" id="header-athlete-btn" onclick="toggleAthleteSwitcher(event)">
          ${MUI.athletes} <span>${escapeHtml(ath.name)}</span>
          <svg class="mui-icon" viewBox="0 0 24 24" style="width:16px;height:16px;"><path d="M7 10l5 5 5-5z"/></svg>
        </button>
        <div class="custom-select-menu" id="header-athlete-menu" style="left:0; min-width:180px;">
          ${athListHtml}
        </div>
      </div>
    `;

    headerActions.innerHTML = `
      <button class="btn-header" onclick="openPreviewModal()" title="Mobile App Preview">
        ${MUI.phone} Preview
      </button>
      <button class="btn-header btn-header-primary" onclick="openExportSessionsModal()" title="Export Standalone HTML">
        ${MUI.export} Export App
      </button>
    `;
    renderAthleteDetail();
  } else if (view === 'builder') {
    const session = ath.pages[activeSessionIndex] || { nav_title: 'Session' };

    // Athlete Switcher dropdown
    const athListHtml = studioAthletes.map(a => `
      <div class="custom-select-item ${a.id === ath.id ? 'selected' : ''}" onclick="navigateTo('builder', '${a.id}', 0); closeAllCustomSelects();">
        <span>${escapeHtml(a.name)}</span>
        ${a.id === ath.id ? MUI.check : ''}
      </div>
    `).join('');

    // Session Switcher dropdown
    const sessListHtml = (ath.pages || []).map((p, pIdx) => `
      <div class="custom-select-item ${pIdx === activeSessionIndex ? 'selected' : ''}" onclick="navigateTo('builder', '${ath.id}', ${pIdx}); closeAllCustomSelects();">
        <span>#${pIdx + 1} ${escapeHtml(p.nav_title || p.title || 'Session')}</span>
        ${pIdx === activeSessionIndex ? MUI.check : ''}
      </div>
    `).join('');

    navBreadcrumb.innerHTML = `
      <div class="header-switcher-wrap">
        <span class="header-switcher-sep">›</span>
        <div style="position:relative; display:inline-block;">
          <button class="switcher-btn" id="header-athlete-btn" onclick="toggleAthleteSwitcher(event)">
            ${MUI.athletes} <span>${escapeHtml(ath.name)}</span>
            <svg class="mui-icon" viewBox="0 0 24 24" style="width:16px;height:16px;"><path d="M7 10l5 5 5-5z"/></svg>
          </button>
          <div class="custom-select-menu" id="header-athlete-menu" style="left:0; min-width:180px;">
            ${athListHtml}
          </div>
        </div>

        <span class="header-switcher-sep">›</span>
        <div style="position:relative; display:inline-block;">
          <button class="switcher-btn" id="header-session-btn" onclick="toggleSessionSwitcher(event)" style="border-color:rgba(0,229,255,0.3); color:var(--teal);">
            ${MUI.dumbbell} <span>${escapeHtml(session.nav_title || session.title || 'Session')}</span>
            <svg class="mui-icon" viewBox="0 0 24 24" style="width:16px;height:16px;"><path d="M7 10l5 5 5-5z"/></svg>
          </button>
          <div class="custom-select-menu" id="header-session-menu" style="left:0; min-width:210px;">
            ${sessListHtml}
          </div>
        </div>
      </div>
    `;

    headerActions.innerHTML = `
      <button class="btn-header icon-only" onclick="navigateTo('athlete')" title="Back to All Sessions">
        ${MUI.arrow_back}
      </button>
      <button class="btn-header" onclick="openPreviewModal()" title="Mobile Preview">
        ${MUI.phone} Preview
      </button>
      <button class="btn-header btn-header-primary" onclick="saveActiveAthleteRoutine()" title="Save Routine Changes">
        ${MUI.save} Save
      </button>
    `;
    renderBuilderView();
  }
}

// 1. ATHLETE ROSTER (HUB) RENDERING
function renderAthletesHub() {
  const grid = document.getElementById('athletes-grid');
  if (!grid) return;

  const countBadge = document.getElementById('athlete-count-badge');
  if (countBadge) countBadge.textContent = `${studioAthletes.length} Athlete${studioAthletes.length === 1 ? '' : 's'}`;

  const query = (document.getElementById('athlete-search-input')?.value || '').toLowerCase().trim();
  const filtered = studioAthletes.filter(a => !query || a.name.toLowerCase().includes(query));

  if (filtered.length === 0) {
    grid.innerHTML = `
      <div style="grid-column: 1 / -1; text-align:center; padding: 60px 20px; color:var(--text-dim);">
        <p style="font-size:1.1em; margin-bottom:12px;">No athletes found matching "${escapeHtml(query)}".</p>
        <button class="btn-header btn-header-primary" onclick="openNewAthleteModal()">${MUI.add} Create Athlete</button>
      </div>
    `;
    return;
  }

  grid.innerHTML = filtered.map(a => {
    const sessionCount = a.pages ? a.pages.length : 0;
    const sessionNames = (a.pages || []).slice(0, 3).map(p => p.nav_title || p.title).join(', ');
    const comments = a.comments || 'No specific notes recorded yet.';

    return `
      <div class="athlete-roster-card" onclick="navigateTo('athlete', '${a.id}')">
        <div>
          <div class="athlete-card-top">
            <div class="athlete-avatar-circle">${escapeHtml(a.initials || a.name.slice(0, 2).toUpperCase())}</div>
            <div>
              <div class="athlete-card-name">${escapeHtml(a.name)}</div>
              <div class="athlete-card-sessions-count">${MUI.dumbbell} ${sessionCount} Workout Session${sessionCount === 1 ? '' : 's'}</div>
            </div>
          </div>

          <div class="athlete-card-comments">
            <strong>Notes:</strong> ${escapeHtml(comments)}
          </div>
          
          <div style="font-size:0.75em; color:var(--text-dim); margin-bottom:16px;">
            <strong>Workouts:</strong> ${escapeHtml(sessionNames || 'None')} ${sessionCount > 3 ? `(+${sessionCount - 3} more)` : ''}
          </div>
        </div>

        <div class="athlete-card-footer" onclick="event.stopPropagation()">
          <button class="btn-header" style="flex:1; justify-content:center;" onclick="navigateTo('athlete', '${a.id}')">
            ${MUI.folder} Open Sessions
          </button>
          <button class="btn-header" title="Export client HTML" onclick="openExportSessionsModal('${a.id}');">${MUI.export}</button>
          <button class="btn-header" title="Delete athlete" style="color:var(--rose);" onclick="deleteAthletePrompt('${a.id}')">${MUI.delete}</button>
        </div>
      </div>
    `;
  }).join('');
}

// 2. ATHLETE DETAIL (HIGH-LEVEL SESSIONS MANAGER)
function renderAthleteDetail() {
  const ath = getActiveAthlete();
  if (!ath) return;

  document.getElementById('detail-athlete-avatar').textContent = ath.initials || ath.name.slice(0, 2).toUpperCase();
  document.getElementById('detail-athlete-name').textContent = ath.name;
  document.getElementById('athlete-comments-input').value = ath.comments || '';
  
  const sessionCount = ath.pages ? ath.pages.length : 0;
  document.getElementById('detail-sessions-count').textContent = `${sessionCount} Session${sessionCount === 1 ? '' : 's'}`;

  const container = document.getElementById('sessions-overview-grid');
  if (!container) return;

  if (!ath.pages || ath.pages.length === 0) {
    container.innerHTML = `
      <div style="text-align:center; padding:50px 20px; color:var(--text-dim); background:var(--bg-card); border-radius:14px; border:1px dashed var(--border);">
        <p style="font-size:1.1em; margin-bottom:12px;">No workout sessions created for ${escapeHtml(ath.name)} yet.</p>
        <button class="btn-header btn-header-primary" onclick="createNewSessionPrompt()">${MUI.add} Add First Session</button>
      </div>
    `;
    return;
  }

  container.innerHTML = ath.pages.map((p, idx) => {
    const cards = p.cards || [];
    let totalSets = 0;
    cards.forEach(c => { totalSets += (c.total_sets || 3); });
    const previewList = cards.slice(0, 4).map(c => c.exercise).join('  -  ');

    return `
      <div class="session-overview-card">
        <div class="soc-left" onclick="navigateTo('builder', '${ath.id}', ${idx})" style="cursor:pointer;">
          <div class="soc-header-row">
            <div class="soc-num">#${idx + 1}</div>
            <div class="soc-title">${escapeHtml(p.title || p.nav_title || 'Session')}</div>
            <span class="soc-badge soc-type-badge">${p.type === 'table' ? MUI.calendar + ' Weekly Plan' : MUI.dumbbell + ' Routine'}</span>
          </div>
          <div class="soc-badges">
            <span class="soc-badge">${cards.length} Exercises</span>
            <span class="soc-badge">${totalSets} Total Sets</span>
          </div>
          <div class="soc-preview-exercises">
            ${escapeHtml(previewList || 'No exercises added yet')} ${cards.length > 4 ? `(+${cards.length - 4} more)` : ''}
          </div>
        </div>

        <div class="soc-actions">
          <button class="btn-header soc-btn-edit" onclick="navigateTo('builder', '${ath.id}', ${idx})">
            ${MUI.edit} Edit Workout Routine
          </button>
          <div class="soc-secondary-actions">
            <button class="btn-header soc-btn-sub" title="Save as Master Template" onclick="saveSessionAsTemplate(${idx})">
              ${MUI.save} <span class="soc-sub-text">Template</span>
            </button>
            <button class="btn-header soc-btn-sub" title="Copy to Another Athlete" onclick="openCopySessionModal(${idx})">
              ${MUI.copy} <span class="soc-sub-text">Copy</span>
            </button>
            <button class="btn-header soc-btn-sub" title="Duplicate Session" onclick="duplicateSessionInAthlete(${idx})">
              ${MUI.duplicate} <span class="soc-sub-text">Duplicate</span>
            </button>
            <button class="btn-header soc-btn-del" title="Delete Session" style="color:var(--rose);" onclick="deleteSessionPrompt(${idx})">
              ${MUI.delete}
            </button>
          </div>
        </div>
      </div>
    `;
  }).join('');

  renderAthleteHistorySection(ath);
}

function saveCurrentAthleteComments(text) {
  const ath = getActiveAthlete();
  if (!ath) return;
  ath.comments = text;
  saveAthletesToStorage();
  const status = document.getElementById('comments-status');
  if (status) {
    status.textContent = 'Saving...';
    setTimeout(() => { status.textContent = 'Auto-saved'; }, 600);
  }
}

// 3. WORKOUT BUILDER / EXERCISE EDITOR
function renderBuilderView() {
  const ath = getActiveAthlete();
  const page = ath.pages[activeSessionIndex];
  if (!page) return;

  document.getElementById('builder-session-title').value = page.title || page.nav_title || 'Session';
  const typeSelect = document.getElementById('builder-session-type');
  if (typeSelect) {
    typeSelect.value = page.type || 'exercises';
    enhanceSelect(typeSelect);
  }

  const container = document.getElementById('exercises-container');
  const addBar = document.getElementById('builder-add-bar');
  if (!container) return;

  if (page.type === 'table') {
    // --- WEEKLY SCHEDULE MODE ---
    if (addBar) {
      addBar.innerHTML = `
        <button class="btn-add-action" onclick="openDayModal(null)">
          ${MUI.add} Add Day to Plan
        </button>
        <button class="btn-add-action" onclick="quickAddRestDay()" style="color:var(--teal); border-color:rgba(0,229,255,0.3);">
          ${MUI.calendar} Quick Add Rest Day
        </button>
        <button class="btn-add-action" onclick="openBatchImportModal()" style="border-color:rgba(167,139,250,0.4); color:var(--purple);">
          ${MUI.copy} Paste from Excel / CSV
        </button>
      `;
    }

    if (!page.cards || page.cards.length === 0) {
      container.innerHTML = `
        <div style="text-align:center; padding: 60px 20px; color: var(--text-dim); background: var(--bg-card); border-radius: 12px; border: 1px dashed var(--border);">
          <div style="font-size:1.15em; font-weight:700; color:var(--text-main); margin-bottom: 6px;">Weekly Training Schedule</div>
          <p style="font-size:0.92em; margin-bottom: 16px;">This schedule has no training days planned yet.</p>
          <button class="btn-header btn-header-primary" onclick="openDayModal(null)">${MUI.add} Add First Day to Plan</button>
        </div>
      `;
      return;
    }

    container.innerHTML = page.cards.map((c, idx) => {
      const dayName = escapeHtml(c.category || `Day ${idx + 1}`);
      const dayFocus = escapeHtml(c.exercise || 'Rest / Recovery');
      const detailsHtml = c.details ? `
        <div class="plan-day-details">${escapeHtml(c.details)}</div>
      ` : '';

      return `
        <div class="exercise-row plan-day-row" draggable="true" data-card-idx="${idx}" onclick="openDayModal(${idx})">
          <div class="row-left">
            <span class="drag-handle" onclick="event.stopPropagation()" style="margin-top:2px;">${MUI.drag}</span>
            <span class="row-num" style="min-width:24px;">#${idx + 1}</span>
            <div style="display:flex; flex-direction:column; gap:4px; flex:1; min-width:0;">
              <div style="display:flex; align-items:center; gap:8px; flex-wrap:wrap;">
                <span class="plan-day-badge">${dayName}</span>
                <span class="plan-day-focus">${dayFocus}</span>
              </div>
              ${detailsHtml}
            </div>
          </div>

          <div class="row-actions" onclick="event.stopPropagation()" style="align-self:flex-start; margin-top:2px;">
            <button type="button" class="btn-row-action edit-btn" onclick="openDayModal(${idx})">
              ${MUI.edit} Edit
            </button>
            <button type="button" class="btn-row-action" title="Duplicate Day" onclick="duplicateDay(${idx})">${MUI.duplicate}</button>
            <button type="button" class="btn-row-action danger" title="Delete" onclick="deleteDayPrompt(${idx})">${MUI.close}</button>
          </div>
        </div>
      `;
    }).join('');

    setupDragAndDrop();
    return;
  }

  // --- EXERCISE LIST MODE ---
  if (addBar) {
    addBar.innerHTML = `
      <button class="btn-add-action" onclick="openExerciseModal(null)">
        ${MUI.add} Add Blank Exercise
      </button>
      <button class="btn-add-action lib-btn" onclick="openLibraryModal()">
        ${MUI.library} Pick from Exercise Library
      </button>
      <button class="btn-add-action" onclick="openBatchImportModal()" style="border-color:rgba(167,139,250,0.4); color:var(--purple);">
        ${MUI.copy} Paste from Excel / CSV
      </button>
    `;
  }

  if (!page.cards || page.cards.length === 0) {
    container.innerHTML = `
      <div style="text-align:center; padding: 60px 20px; color: var(--text-dim); background: var(--bg-card); border-radius: 12px; border: 1px dashed var(--border);">
        <p style="font-size:1.05em; margin-bottom: 12px;">This session has no exercises yet.</p>
        <button class="btn-header btn-header-primary" onclick="openLibraryModal()">${MUI.library} Pick from Exercise Library</button>
      </div>
    `;
    return;
  }

  container.innerHTML = page.cards.map((c, idx) => {
    const cueHtml = c.cue ? `<div class="row-cue-banner" onclick="event.stopPropagation()"><span style="color:var(--purple); font-weight:700;">Cue:</span> "${escapeHtml(c.cue)}"</div>` : '';
    const supersetBadge = c.superset ? `<span class="row-badge" style="background:rgba(0, 229, 255, 0.12); color:var(--teal); font-weight:800; padding:2px 6px;">${escapeHtml(c.superset)}</span>` : '';

    return `
    <div class="exercise-row" draggable="true" data-card-idx="${idx}" onclick="openExerciseModal(${idx})">
      <div class="row-left">
        <span class="drag-handle" onclick="event.stopPropagation()">${MUI.drag}</span>
        <span class="row-num">#${idx + 1}</span>
        ${supersetBadge}
        <span class="row-cat-badge">${escapeHtml(c.category || 'Exercise')}</span>
        <span class="row-title">${escapeHtml(c.exercise || 'Unnamed')}</span>
      </div>

      <div class="row-pill-info" onclick="event.stopPropagation()">
        <span class="row-badge">${MUI.trending_up} <span class="badge-val">${escapeHtml(c.details || '3 x 10')}</span></span>
        <span class="row-badge">${MUI.dumbbell} <span class="badge-val">${escapeHtml(c.weight || 'BW')}</span></span>
        <span class="row-badge">${MUI.timer} <span class="badge-val">${escapeHtml(c.rest || '60s')}</span></span>
      </div>

      ${cueHtml}

      <div class="row-actions" onclick="event.stopPropagation()">
        <button type="button" class="btn-row-action edit-btn" onclick="openExerciseModal(${idx})">
          ${MUI.edit} Edit
        </button>
        <button type="button" class="btn-row-action" title="Duplicate Exercise" onclick="duplicateExercise(${idx})">${MUI.duplicate}</button>
        <button type="button" class="btn-row-action danger" title="Delete" onclick="deleteExercise(${idx})">${MUI.close}</button>
      </div>
    </div>
  `;
  }).join('');

  setupDragAndDrop();
}

function updateActiveSessionTitle(title) {
  const ath = getActiveAthlete();
  const page = ath.pages[activeSessionIndex];
  if (page) {
    page.title = title;
    page.nav_title = title;
    saveAthletesToStorage();
    showToast(` Updated title to "${title}"`);
  }
}

function updateActiveSessionType(type) {
  const ath = getActiveAthlete();
  const page = ath.pages[activeSessionIndex];
  if (page) {
    page.type = type;
    saveAthletesToStorage();
    renderBuilderView();
    showToast(`Switched to ${type === 'table' ? 'Weekly Plan Schedule' : 'Exercise List Routine'}`);
  }
}


// DAY MODAL CRUD ACTIONS (WEEKLY SCHEDULE)
let modalEditingDayIdx = null;

function quickFillDayField(id, val) {
  const el = document.getElementById(id);
  if (el) el.value = val;
}

function openDayModal(dayIdx) {
  modalEditingDayIdx = dayIdx;
  const ath = getActiveAthlete();
  const page = ath.pages[activeSessionIndex];

  if (dayIdx === null || dayIdx === undefined) {
    document.getElementById('day-modal-title').innerHTML = `${MUI.add} Add Schedule Day`;
    const count = (page.cards || []).length;
    const days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
    const defaultDay = count < days.length ? days[count] : `Day ${count + 1}`;
    document.getElementById('modal-day-name').value = defaultDay;
    document.getElementById('modal-day-focus').value = '';
    document.getElementById('modal-day-details').value = '';
  } else {
    const card = page.cards[dayIdx];
    document.getElementById('day-modal-title').innerHTML = `${MUI.edit} Edit Schedule Day`;
    document.getElementById('modal-day-name').value = card.category || `Day ${dayIdx + 1}`;
    document.getElementById('modal-day-focus').value = card.exercise || '';
    document.getElementById('modal-day-details').value = card.details || '';
  }

  openModal('day-modal');
}

function saveDayModal() {
  const name = document.getElementById('modal-day-name').value.trim();
  if (!name) {
    showToast('Please enter a day label (e.g. Monday)');
    return;
  }
  const focus = document.getElementById('modal-day-focus').value.trim();
  const details = document.getElementById('modal-day-details').value.trim();

  const ath = getActiveAthlete();
  const page = ath.pages[activeSessionIndex];
  if (!page.cards) page.cards = [];

  const dayData = {
    category: name,
    exercise: focus,
    details: details,
    total_sets: 0
  };

  if (modalEditingDayIdx === null || modalEditingDayIdx === undefined) {
    page.cards.push(dayData);
    showToast(`Added ${name} to schedule`);
  } else {
    page.cards[modalEditingDayIdx] = { ...page.cards[modalEditingDayIdx], ...dayData };
    showToast(`Updated ${name}`);
  }

  saveAthletesToStorage();
  closeModal('day-modal');
  renderBuilderView();
}

function duplicateDay(idx) {
  const ath = getActiveAthlete();
  const page = ath.pages[activeSessionIndex];
  if (!page || !page.cards || !page.cards[idx]) return;
  const clone = JSON.parse(JSON.stringify(page.cards[idx]));
  clone.category = `${clone.category} (Copy)`;
  page.cards.splice(idx + 1, 0, clone);
  saveAthletesToStorage();
  renderBuilderView();
  showToast('Day duplicated');
}

function deleteDayPrompt(idx) {
  const ath = getActiveAthlete();
  const page = ath.pages[activeSessionIndex];
  if (!page || !page.cards || !page.cards[idx]) return;
  const dayName = page.cards[idx].category || `Day #${idx + 1}`;
  showConfirmModal('Delete Day', `Are you sure you want to remove ${dayName} from this weekly schedule?`, () => {
    page.cards.splice(idx, 1);
    saveAthletesToStorage();
    renderBuilderView();
    showToast('Day removed from plan');
  });
}

function quickAddRestDay() {
  const ath = getActiveAthlete();
  const page = ath.pages[activeSessionIndex];
  if (!page) return;
  if (!page.cards) page.cards = [];
  const count = page.cards.length;
  const days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
  const defaultDay = count < days.length ? days[count] : `Day ${count + 1}`;
  page.cards.push({
    category: defaultDay,
    exercise: 'Rest & Recovery',
    details: 'Full rest day or 20-30 mins light mobility / walking.',
    total_sets: 0
  });
  saveAthletesToStorage();
  renderBuilderView();
  showToast(`Added Rest Day (${defaultDay})`);
}

function saveActiveAthleteRoutine() {
  saveAthletesToStorage();
  showToast('Routine Saved!', MUI.check);
}

// SESSION CRUD ACTIONS
function createNewSessionPrompt() {
  const ath = getActiveAthlete();
  const nextNum = (ath.pages ? ath.pages.length : 0) + 1;
  
  showPromptModal('Add New Session', 'Session Title', `Session ${nextNum} - Workout`, (title) => {
    if (!title) return;
    if (!ath.pages) ath.pages = [];
    
    ath.pages.push({
      id: `page-custom-${Date.now()}`,
      title: title,
      nav_title: title,
      type: 'exercises',
      cards: []
    });

    saveAthletesToStorage();
    showToast(` Created "${title}"!`);
    navigateTo('athlete');
  });
}

function duplicateSessionInAthlete(sessionIdx) {
  const ath = getActiveAthlete();
  const session = ath.pages[sessionIdx];
  if (!session) return;

  const cloned = JSON.parse(JSON.stringify(session));
  cloned.id = `page-cloned-${Date.now()}`;
  cloned.title = `${session.title || session.nav_title} (Copy)`;
  cloned.nav_title = `${session.nav_title || session.title} (Copy)`;

  ath.pages.splice(sessionIdx + 1, 0, cloned);
  saveAthletesToStorage();
  showToast(` Duplicated "${session.nav_title || session.title}"!`);
  renderAthleteDetail();
}

function deleteSessionPrompt(sessionIdx) {
  const ath = getActiveAthlete();
  const session = ath.pages[sessionIdx];
  if (!session) return;

  showConfirmModal(
    'Delete Session',
    `Are you sure you want to delete "${session.nav_title || session.title}"? This cannot be undone.`,
    (confirmed) => {
      if (confirmed) {
        ath.pages.splice(sessionIdx, 1);
        saveAthletesToStorage();
        showToast(` Deleted session.`);
        renderAthleteDetail();
      }
    }
  );
}

// 1-CLICK COPY SESSION TO ANOTHER ATHLETE
function openCopySessionModal(sessionIdx) {
  const ath = getActiveAthlete();
  const session = ath.pages[sessionIdx];
  if (!session) return;

  sessionToCopyIdx = sessionIdx;
  document.getElementById('copy-session-name').textContent = `"${session.nav_title || session.title}"`;

  const select = document.getElementById('copy-destination-athlete');
  select.innerHTML = studioAthletes
    .filter(a => a.id !== ath.id)
    .map(a => `<option value="${a.id}">${escapeHtml(a.name)} (${a.pages ? a.pages.length : 0} sessions)</option>`)
    .join('');

  if (select.options.length === 0) {
    showToast('No other athletes available to copy to. Create another athlete first!', MUI.warning);
    return;
  }

  const copySelect = document.getElementById('copy-destination-athlete');
  if (copySelect) {
    copySelect.dataset.customized = 'false';
    const oldWrap = copySelect.closest('.custom-select-wrapper');
    if (oldWrap) {
      oldWrap.parentNode.insertBefore(copySelect, oldWrap);
      oldWrap.remove();
    }
    enhanceSelect(copySelect);
  }
  openModal('copy-session-modal');
}

function confirmCopySessionToAthlete() {
  const destId = document.getElementById('copy-destination-athlete').value;
  const destAthlete = studioAthletes.find(a => a.id === destId);
  const srcAthlete = getActiveAthlete();
  const session = srcAthlete.pages[sessionToCopyIdx];

  if (!destAthlete || !session) return;

  const cloned = JSON.parse(JSON.stringify(session));
  cloned.id = `page-copied-${Date.now()}`;
  if (!destAthlete.pages) destAthlete.pages = [];
  destAthlete.pages.push(cloned);

  saveAthletesToStorage();
  closeModal('copy-session-modal');
  showToast(` Copied "${session.nav_title || session.title}" to ${destAthlete.name}!`);
}

// MASTER SESSION TEMPLATES CONTROLLER
function renderTemplatesView() {
  const grid = document.getElementById('templates-grid');
  if (!grid) return;

  const countBadge = document.getElementById('template-count-badge');
  if (countBadge) countBadge.textContent = `${sessionTemplates.length} Template${sessionTemplates.length === 1 ? '' : 's'}`;

  const query = (document.getElementById('template-search-input')?.value || '').toLowerCase().trim();
  const list = sessionTemplates.filter(t => {
    return !query || t.name.toLowerCase().includes(query) || (t.category && t.category.toLowerCase().includes(query));
  });

  if (list.length === 0) {
    grid.innerHTML = `
      <div style="grid-column: 1 / -1; text-align:center; padding: 60px 20px; color: var(--text-dim); background: var(--bg-panel); border-radius: var(--radius); border: 1px dashed var(--border);">
        <p style="font-size:1.1em; margin-bottom: 12px; font-weight:700;">No session templates found.</p>
        <button class="btn-header btn-header-primary" onclick="openCreateTemplateModal()">${MUI.add} Create New Template</button>
      </div>
    `;
    return;
  }

  grid.innerHTML = list.map(t => {
    const cards = t.cards || [];
    let totalSets = 0;
    cards.forEach(c => { totalSets += (c.total_sets || 3); });
    const previewList = cards.map(c => c.exercise).join('  -  ');

    return `
      <div class="athlete-roster-card" style="cursor:default;">
        <div>
          <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:12px;">
            <div>
              <div style="font-size:0.7em; font-weight:800; text-transform:uppercase; color:var(--purple); margin-bottom:4px;">
                ${escapeHtml(t.category || 'WORKOUT TEMPLATE')}
              </div>
              <div style="font-size:1.15em; font-weight:800; color:var(--text-main);">
                ${escapeHtml(t.name)}
              </div>
              <div style="font-size:0.78em; color:var(--text-muted); margin-top:3px;">
                ${cards.length} Exercises  -  ${totalSets} Total Sets
              </div>
            </div>
            <div style="display:flex; gap:6px;">
              <button class="btn-header" style="padding:4px 8px; font-size:0.76em;" title="Rename Template" onclick="renameTemplate('${t.id}')">${MUI.edit}</button>
              <button class="btn-header" style="padding:4px 8px; font-size:0.76em; color:var(--rose);" title="Delete Template" onclick="deleteTemplate('${t.id}')">${MUI.delete}</button>
            </div>
          </div>

          <div style="background:rgba(0,0,0,0.25); border:1px solid rgba(255,255,255,0.04); border-radius:10px; padding:10px 12px; margin-bottom:16px; font-size:0.78em; color:var(--text-dim); line-height:1.5; max-height:80px; overflow-y:auto;">
            ${escapeHtml(previewList || 'No exercises')}
          </div>
        </div>

        <button class="btn-header btn-header-primary" onclick="openAssignTemplateModal('${t.id}')" style="width:100%; justify-content:center; padding:10px 16px;">
          ${MUI.rocket} Assign to Athletes...
        </button>
      </div>
    `;
  }).join('');
}

let currentAssigningTemplateId = null;

function openAssignTemplateModal(tplId) {
  currentAssigningTemplateId = tplId;
  const tpl = sessionTemplates.find(t => t.id === tplId);
  if (!tpl) return;

  const titleEl = document.getElementById('assign-template-title');
  const nameInput = document.getElementById('assign-custom-session-name');
  const listEl = document.getElementById('assign-athletes-list');

  if (titleEl) titleEl.innerText = `Assign "${tpl.name}"`;
  if (nameInput) nameInput.value = tpl.name;

  if (listEl) {
    if (studioAthletes.length === 0) {
      listEl.innerHTML = '<div style="color:var(--text-dim); padding:12px; text-align:center;">No athletes found. Create an athlete first.</div>';
    } else {
      listEl.innerHTML = studioAthletes.map(ath => `
        <label style="display:flex; align-items:center; gap:12px; background:var(--bg-card); border:1px solid var(--border); border-radius:10px; padding:10px 14px; cursor:pointer; transition:0.15s all;"
               onmouseover="this.style.borderColor='rgba(0,229,255,0.3)'" onmouseout="this.style.borderColor='var(--border)'">
          <input type="checkbox" class="assign-ath-checkbox" value="${ath.id}" checked style="width:18px; height:18px; cursor:pointer;">
          <div class="athlete-avatar-circle" style="width:34px; height:34px; font-size:0.85em; border-radius:8px;">${escapeHtml(ath.initials || ath.name.slice(0,2).toUpperCase())}</div>
          <div style="flex:1;">
            <div style="font-weight:700; color:var(--text-main); font-size:0.9em;">${escapeHtml(ath.name)}</div>
            <div style="font-size:0.75em; color:var(--text-muted);">${(ath.pages || []).length} current sessions</div>
          </div>
        </label>
      `).join('');
    }
  }

  openModal('assign-template-modal');
}

function toggleSelectAllAssignAthletes() {
  const boxes = document.querySelectorAll('.assign-ath-checkbox');
  const allChecked = Array.from(boxes).every(b => b.checked);
  boxes.forEach(b => b.checked = !allChecked);
}

function confirmAssignTemplateToAthletes() {
  const tpl = sessionTemplates.find(t => t.id === currentAssigningTemplateId);
  if (!tpl) return;

  const customName = (document.getElementById('assign-custom-session-name')?.value || '').trim() || tpl.name;
  const checkedBoxes = document.querySelectorAll('.assign-ath-checkbox:checked');
  const selectedIds = Array.from(checkedBoxes).map(b => b.value);

  if (selectedIds.length === 0) {
    showToast(' Please select at least one athlete.');
    return;
  }

  const assignedNames = [];
  selectedIds.forEach(athId => {
    const athlete = studioAthletes.find(a => a.id === athId);
    if (athlete) {
      if (!athlete.pages) athlete.pages = [];
      athlete.pages.push({
        id: `page-tpl-${Date.now()}-${Math.random().toString(36).substr(2, 4)}`,
        title: customName,
        nav_title: customName,
        type: 'exercises',
        cards: JSON.parse(JSON.stringify(tpl.cards || []))
      });
      assignedNames.push(athlete.name);
    }
  });

  saveAthletesToStorage();
  closeModal('assign-template-modal');
  showToast(` Assigned "${customName}" to ${assignedNames.length} athlete${assignedNames.length === 1 ? '' : 's'}!`);

  if (currentView === 'athlete') renderAthleteDetail();
  else if (currentView === 'hub') renderAthletesHub();
}

function openCreateTemplateModal() {
  const tplCat = document.getElementById('template-category-input');
  if (tplCat) enhanceSelect(tplCat);
  document.getElementById('template-name-input').value = '';
  document.getElementById('template-category-input').value = 'Push Focus';
  openModal('create-template-modal');
}

function confirmCreateTemplate() {
  const name = (document.getElementById('template-name-input')?.value || '').trim();
  if (!name) {
    showToast(' Please enter a template name.');
    return;
  }
  const category = document.getElementById('template-category-input')?.value || 'Custom Focus';

  const newTpl = {
    id: `tpl_${Date.now()}`,
    name: name,
    category: category,
    cards: []
  };

  sessionTemplates.unshift(newTpl);
  saveTemplatesToStorage();
  closeModal('create-template-modal');
  showToast(` Created session template "${name}"!`);
  renderTemplatesView();
}

function renameTemplate(tplId) {
  const tpl = sessionTemplates.find(t => t.id === tplId);
  if (!tpl) return;
  showPromptModal('Rename Template', 'Template Name', tpl.name, (newName) => {
    if (!newName) return;
    tpl.name = newName;
    saveTemplatesToStorage();
    showToast(` Renamed template to "${newName}"`);
    renderTemplatesView();
  });
}

function deleteTemplate(tplId) {
  const tpl = sessionTemplates.find(t => t.id === tplId);
  if (!tpl) return;
  showConfirmModal('Delete Template', `Delete the "${tpl.name}" session template?`, (confirmed) => {
    if (confirmed) {
      sessionTemplates = sessionTemplates.filter(t => t.id !== tplId);
      saveTemplatesToStorage();
      showToast(' Deleted template.');
      renderTemplatesView();
    }
  });
}

function saveSessionAsTemplate(sessionIdx) {
  const ath = getActiveAthlete();
  if (!ath || !ath.pages || !ath.pages[sessionIdx]) return;
  const session = ath.pages[sessionIdx];
  const defaultName = session.title || session.nav_title || 'Workout Session';

  showPromptModal('Save Session as Template', 'Template Name', defaultName, (templateName) => {
    if (!templateName) return;
    const newTpl = {
      id: `tpl_saved_${Date.now()}`,
      name: templateName,
      category: `${ath.name.split(' ')[0]}'s Routine`,
      cards: JSON.parse(JSON.stringify(session.cards || []))
    };
    sessionTemplates.unshift(newTpl);
    saveTemplatesToStorage();
    showToast(` Saved "${templateName}" to Master Session Templates!`);
  });
}

// EXERCISE CRUD ACTIONS
function openExerciseModal(cardIdx) {
  modalEditingCardIdx = cardIdx;
  const ath = getActiveAthlete();
  const page = ath.pages[activeSessionIndex];

  const catSelect = document.getElementById('modal-ex-category');
  if (catSelect) enhanceSelect(catSelect);
  if (cardIdx === null || cardIdx === undefined) {
    document.getElementById('exercise-modal-title').textContent = '${MUI.add} Add New Exercise';
    document.getElementById('modal-ex-category').value = 'Primary Strength';
    document.getElementById('modal-ex-name').value = '';
    document.getElementById('modal-ex-details').value = '3 x 10';
    document.getElementById('modal-ex-weight').value = '';
    document.getElementById('modal-ex-rest').value = '60s';
    document.getElementById('modal-ex-cue').value = '';
    document.getElementById('modal-ex-videos').value = '';
  } else {
    const card = page.cards[cardIdx];
    document.getElementById('exercise-modal-title').textContent = '${MUI.edit} Edit Exercise Details';
    document.getElementById('modal-ex-category').value = card.category || 'Primary Strength';
    document.getElementById('modal-ex-name').value = card.exercise || '';
    document.getElementById('modal-ex-details').value = card.details || '';
    document.getElementById('modal-ex-weight').value = card.weight || '';
    document.getElementById('modal-ex-rest').value = card.rest || '60s';
    if (document.getElementById('modal-ex-superset')) document.getElementById('modal-ex-superset').value = card.superset || '';
    document.getElementById('modal-ex-cue').value = card.cue || '';
    document.getElementById('modal-ex-videos').value = (card.videos || []).join('\\n');
  }

  openModal('exercise-modal');
}

function saveExerciseModal() {
  const name = document.getElementById('modal-ex-name').value.trim();
  if (!name) {
    showToast(' Please enter an exercise name', '');
    return;
  }

  const ath = getActiveAthlete();
  const page = ath.pages[activeSessionIndex];
  if (!page.cards) page.cards = [];

  const details = document.getElementById('modal-ex-details').value.trim() || '3 x 10';
  let totalSets = 3;
  const match = details.match(/^(\\d+)\\s*[xX]/);
  if (match) totalSets = parseInt(match[1]);

  const restStr = document.getElementById('modal-ex-rest').value.trim() || '60s';
  let restSec = 60;
  const restMatch = restStr.match(/(\\d+)/);
  if (restMatch) restSec = parseInt(restMatch[1]);

  const videos = document.getElementById('modal-ex-videos').value
    .split('\\n')
    .map(v => v.trim())
    .filter(v => v.length > 0);

  const cardData = {
    category: document.getElementById('modal-ex-category').value,
    exercise: name,
    details: details,
    weight: document.getElementById('modal-ex-weight').value.trim(),
    rest: restStr,
    rest_seconds: restSec,
    total_sets: totalSets,
    cue: document.getElementById('modal-ex-cue').value.trim(),
    videos: videos
  };

  if (modalEditingCardIdx === null) {
    page.cards.push(cardData);
    showToast(` Added "${name}"`);
  } else {
    page.cards[modalEditingCardIdx] = cardData;
    showToast(` Updated "${name}"`);
  }

  closeModal('exercise-modal');
  saveAthletesToStorage();
  renderBuilderView();
}

function duplicateExercise(cardIdx) {
  const ath = getActiveAthlete();
  const page = ath.pages[activeSessionIndex];
  const card = page.cards[cardIdx];
  if (!card) return;

  const cloned = JSON.parse(JSON.stringify(card));
  page.cards.splice(cardIdx + 1, 0, cloned);
  saveAthletesToStorage();
  showToast(` Duplicated "${card.exercise}"`);
  renderBuilderView();
}

function deleteExercise(cardIdx) {
  const ath = getActiveAthlete();
  const page = ath.pages[activeSessionIndex];
  const card = page.cards[cardIdx];
  if (!card) return;

  page.cards.splice(cardIdx, 1);
  saveAthletesToStorage();
  showToast(` Removed "${card.exercise}"`);
  renderBuilderView();
}

// DRAG AND DROP REORDERING
function setupDragAndDrop() {
  const container = document.getElementById('exercises-container');
  if (!container) return;

  const rows = container.querySelectorAll('.exercise-row');
  let draggedEl = null;

  rows.forEach(row => {
    row.addEventListener('dragstart', (e) => {
      draggedEl = row;
      row.classList.add('dragging');
      e.dataTransfer.effectAllowed = 'move';
    });
    row.addEventListener('dragend', () => {
      row.classList.remove('dragging');
      const newRows = [...container.querySelectorAll('.exercise-row')];
      const ath = getActiveAthlete();
      const page = ath.pages[activeSessionIndex];
      const newCards = [];
      newRows.forEach(r => {
        const oldIdx = parseInt(r.getAttribute('data-card-idx'));
        if (page.cards[oldIdx]) newCards.push(page.cards[oldIdx]);
      });
      page.cards = newCards;
      saveAthletesToStorage();
      renderBuilderView();
    });
    row.addEventListener('dragover', (e) => {
      e.preventDefault();
      const afterElement = getDragAfterElement(container, e.clientY);
      if (afterElement == null) {
        container.appendChild(draggedEl);
      } else {
        container.insertBefore(draggedEl, afterElement);
      }
    });
  });
}

function getDragAfterElement(container, y) {
  const draggableElements = [...container.querySelectorAll('.exercise-row:not(.dragging)')];
  return draggableElements.reduce((closest, child) => {
    const box = child.getBoundingClientRect();
    const offset = y - box.top - box.height / 2;
    if (offset < 0 && offset > closest.offset) {
      return { offset: offset, element: child };
    } else {
      return closest;
    }
  }, { offset: Number.NEGATIVE_INFINITY }).element;
}

// ATHLETE CREATION & DELETION
function openNewAthleteModal() {
  document.getElementById('new-ath-name').value = '';
  document.getElementById('new-ath-comments').value = '';
  openModal('new-athlete-modal');
}

function confirmCreateAthlete() {
  const name = document.getElementById('new-ath-name').value.trim();
  if (!name) {
    showToast(' Please provide an athlete full name.', '');
    return;
  }

  const templateChoice = document.getElementById('new-ath-template').value;
  let pages = [];

  if (templateChoice === 'jada') {
    const jada = studioAthletes.find(a => a.id === 'ath_1') || window.INITIAL_ATHLETES[0];
    pages = JSON.parse(JSON.stringify(jada.pages || []));
  } else if (templateChoice === 'marcus') {
    const marcus = studioAthletes.find(a => a.id === 'ath_2') || window.INITIAL_ATHLETES[1];
    pages = JSON.parse(JSON.stringify(marcus.pages || []));
  } else {
    pages = [{
      id: `page-fresh-${Date.now()}`,
      title: 'Session 1 - Foundation',
      nav_title: 'Session 1',
      type: 'exercises',
      cards: []
    }];
  }

  const initials = name.split(' ').map(p => p[0]).join('').toUpperCase().slice(0, 2);
  const newAthlete = {
    id: `ath_${Date.now()}`,
    name: name,
    initials: initials,
    comments: document.getElementById('new-ath-comments').value.trim(),
    pages: pages,
    logs: []
  };

  studioAthletes.push(newAthlete);
  saveAthletesToStorage();
  closeModal('new-athlete-modal');
  showToast(` Athlete profile for "${name}" created!`);
  navigateTo('athlete', newAthlete.id);
}

function deleteAthletePrompt(athleteId) {
  const ath = studioAthletes.find(a => a.id === athleteId);
  if (!ath) return;

  showConfirmModal(
    'Delete Athlete Profile',
    `Are you sure you want to delete "${ath.name}" and all their training programs? This cannot be undone.`,
    (confirmed) => {
      if (confirmed) {
        studioAthletes = studioAthletes.filter(a => a.id !== athleteId);
        if (studioAthletes.length === 0) {
          studioAthletes = window.INITIAL_ATHLETES || [];
        }
        activeAthleteId = studioAthletes[0].id;
        saveAthletesToStorage();
        showToast(` Deleted "${ath.name}".`);
        navigateTo('hub');
      }
    }
  );
}

// EXERCISE LIBRARY & CUSTOM EXERCISES
let activeLibraryCategory = 'All';
let customLibrary = [];
try {
  const savedLib = localStorage.getItem('coach_studio_custom_library');
  if (savedLib) customLibrary = JSON.parse(savedLib);
} catch (e) {}

function getAllLibraryExercises() {
  return [...DEFAULT_LIBRARY, ...customLibrary];
}

function openLibraryModal() {
  renderLibraryPills();
  renderLibraryCards();
  openModal('library-modal');
}

function renderLibraryPills() {
  const container = document.getElementById('lib-pills');
  if (!container) return;
  const groups = ['All', 'Warm-Up', 'Upper Body', 'Lower Body', 'Core'];
  container.innerHTML = groups.map(g => `
    <div class="lib-filter-pill ${activeLibraryCategory === g ? 'active' : ''}" onclick="setLibraryFilter('${g}')">
      ${g}
    </div>
  `).join('');
}

function setLibraryFilter(group) {
  activeLibraryCategory = group;
  renderLibraryPills();
  renderLibraryCards();
}

function renderLibraryCards() {
  const container = document.getElementById('lib-cards');
  if (!container) return;

  const query = (document.getElementById('lib-search')?.value || '').toLowerCase().trim();
  const allList = getAllLibraryExercises();
  const list = allList.filter((item, origIdx) => {
    item._origIdx = origIdx;
    const matchCat = (activeLibraryCategory === 'All') || (item.group === activeLibraryCategory) || (item.category === activeLibraryCategory);
    const matchQ = !query || item.exercise.toLowerCase().includes(query) || (item.category && item.category.toLowerCase().includes(query));
    return matchCat && matchQ;
  });

  if (list.length === 0) {
    container.innerHTML = `<div style="grid-column: 1 / -1; text-align:center; padding: 40px; color: var(--text-dim);">No exercises found.</div>`;
    return;
  }

  // Determine button text and action based on user's current view context
  let btnLabel = `${MUI.add} Add to Session`;
  let btnActionType = 'session';

  if (currentView === 'builder') {
    const ath = getActiveAthlete();
    const curSess = (ath && ath.pages && ath.pages[activeSessionIndex]) ? (ath.pages[activeSessionIndex].nav_title || ath.pages[activeSessionIndex].title || 'Session') : 'Session';
    btnLabel = `${MUI.add} Add to "${curSess}"`;
    btnActionType = 'session';
  } else if (currentView === 'athlete') {
    const ath = getActiveAthlete();
    btnLabel = `${MUI.add} Add to ${ath ? ath.name.split(' ')[0] : 'Athlete'}...`;
    btnActionType = 'athlete';
  } else {
    btnLabel = `${MUI.add} Add to Athlete...`;
    btnActionType = 'hub';
  }

  container.innerHTML = list.map(item => `
    <div class="lib-item-card">
      <div>
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:4px;">
          <div style="font-size:0.7em; font-weight:800; text-transform:uppercase; color:var(--purple);">${escapeHtml(item.category || item.group)}</div>
          ${item.isCustom ? '<span style="font-size:0.65em; background:rgba(0,229,255,0.15); color:var(--teal); padding:2px 6px; border-radius:6px; font-weight:700;">CUSTOM</span>' : ''}
        </div>
        <div style="font-size:0.96em; font-weight:800; color:var(--text-main); margin-bottom:4px;">${escapeHtml(item.exercise)}</div>
        <div style="font-size:0.78em; color:var(--text-muted);">${escapeHtml(item.details)}  -  ${escapeHtml(item.weight || 'Bodyweight')}  -  ${escapeHtml(item.rest || '60s')}</div>
        ${item.cue ? `<div style="font-size:0.74em; color:var(--text-dim); margin-top:4px; font-style:italic;">${escapeHtml(item.cue)}</div>` : ''}
      </div>
      <button class="btn-lib-insert" onclick="handleLibraryAction(${item._origIdx}, '${btnActionType}', this)">
        ${btnLabel}
      </button>
    </div>
  `).join('');
}

let pendingLibraryExercise = null;

function handleLibraryAction(origIdx, actionType, btn) {
  const allList = getAllLibraryExercises();
  const selected = allList[origIdx];
  if (!selected) return;
  pendingLibraryExercise = selected;

  if (actionType === 'session') {
    insertLibraryExercise(selected, btn);
  } else if (actionType === 'athlete') {
    const ath = getActiveAthlete();
    if (!ath || !ath.pages || ath.pages.length === 0) {
      showToast(' Athlete has no sessions yet! Add a session first.');
      return;
    }
    openAthleteSessionPickerModal(ath, selected);
  } else {
    openGlobalAthleteSessionPickerModal(selected);
  }
}

function insertLibraryExercise(selected, btn) {
  const ath = getActiveAthlete();
  if (!ath) return;
  const page = ath.pages[activeSessionIndex];
  if (!page) return;
  if (!page.cards) page.cards = [];

  page.cards.push({
    category: selected.category || 'Primary Strength',
    exercise: selected.exercise,
    details: selected.details,
    weight: selected.weight || '',
    rest: selected.rest || '60s',
    rest_seconds: selected.rest_seconds || 60,
    total_sets: selected.total_sets || 3,
    cue: selected.cue || '',
    videos: selected.videos ? [...selected.videos] : []
  });

  if (btn) {
    btn.innerHTML = MUI.check + ' Added to Session!';
    btn.style.background = 'linear-gradient(135deg, var(--teal), #00A3FF)';
    btn.style.color = '#000';
    btn.style.fontWeight = '800';
    btn.style.borderColor = 'transparent';
  }

  saveAthletesToStorage();
  showToast(` Added "${selected.exercise}" to session!`);

  setTimeout(() => {
    closeModal('library-modal');
    if (currentView === 'builder') renderBuilderView();
    else if (currentView === 'athlete') renderAthleteDetail();
  }, 320);
}

// ATHLETE SESSION PICKER MODAL
function openAthleteSessionPickerModal(ath, exercise) {
  const titleEl = document.getElementById('pick-session-title');
  const subEl = document.getElementById('pick-session-subtitle');
  const listEl = document.getElementById('pick-session-list');

  if (titleEl) titleEl.innerText = `Add "${exercise.exercise}"`;
  if (subEl) subEl.innerText = `Select session in ${ath.name}'s program:`;

  if (listEl) {
    listEl.innerHTML = ath.pages.map((p, pIdx) => `
      <div style="background:var(--bg-card); border:1.5px solid var(--border); border-radius:12px; padding:12px 16px; cursor:pointer; display:flex; justify-content:space-between; align-items:center; transition:0.18s all;"
           onmouseover="this.style.borderColor='var(--teal)'; this.style.boxShadow='0 0 12px var(--teal-glow)';"
           onmouseout="this.style.borderColor='var(--border)'; this.style.boxShadow='none';"
           onclick="addExerciseToAthleteSessionIndex(${pIdx})">
        <div>
          <div style="font-size:0.94em; font-weight:800; color:var(--text-main); margin-bottom:3px;">
            ${escapeHtml(p.title || p.nav_title || `Session ${pIdx+1}`)}
          </div>
          <div style="font-size:0.75em; color:var(--text-muted);">
            ${p.type === 'table' ? MUI.calendar + ' Weekly Plan' : `${MUI.dumbbell} ${(p.cards || []).length} Exercises`}
          </div>
        </div>
        <div style="font-size:1.1em; color:var(--teal); font-weight:800;">${MUI.add}</div>
      </div>
    `).join('');
  }

  openModal('pick-session-modal');
}

function addExerciseToAthleteSessionIndex(pIdx) {
  const ath = getActiveAthlete();
  if (!ath || !pendingLibraryExercise) return;
  const page = ath.pages[pIdx];
  if (!page) return;
  if (!page.cards) page.cards = [];

  page.cards.push({
    category: pendingLibraryExercise.category || 'Primary Strength',
    exercise: pendingLibraryExercise.exercise,
    details: pendingLibraryExercise.details,
    weight: pendingLibraryExercise.weight || '',
    rest: pendingLibraryExercise.rest || '60s',
    rest_seconds: pendingLibraryExercise.rest_seconds || 60,
    total_sets: pendingLibraryExercise.total_sets || 3,
    cue: pendingLibraryExercise.cue || '',
    videos: pendingLibraryExercise.videos ? [...pendingLibraryExercise.videos] : []
  });

  saveAthletesToStorage();
  closeModal('pick-session-modal');
  closeModal('library-modal');
  showToast(` Added "${pendingLibraryExercise.exercise}" to ${page.title || 'session'}!`);

  if (currentView === 'athlete') renderAthleteDetail();
  else if (currentView === 'builder') renderBuilderView();
}

// GLOBAL ATHLETE & SESSION PICKER
function openGlobalAthleteSessionPickerModal(exercise) {
  const titleEl = document.getElementById('pick-global-title');
  const subEl = document.getElementById('pick-global-subtitle');
  if (titleEl) titleEl.innerText = `Add "${exercise.exercise}"`;
  if (subEl) subEl.innerText = `Select destination athlete and session:`;

  const athSelect = document.getElementById('pick-global-ath-select');
  if (athSelect) {
    athSelect.innerHTML = studioAthletes.map(a => `
      <option value="${a.id}">${escapeHtml(a.name)}</option>
    `).join('');
  }

  updateGlobalSessionPickerOptions();
  openModal('pick-global-modal');
}

function updateGlobalSessionPickerOptions() {
  const athSelect = document.getElementById('pick-global-ath-select');
  const sessSelect = document.getElementById('pick-global-sess-select');
  if (!athSelect || !sessSelect) return;

  const athId = athSelect.value;
  const athlete = studioAthletes.find(a => a.id === athId);
  if (!athlete || !athlete.pages || athlete.pages.length === 0) {
    sessSelect.innerHTML = '<option value="">(No sessions available)</option>';
    return;
  }

  sessSelect.innerHTML = athlete.pages.map((p, idx) => `
    <option value="${idx}">${escapeHtml(p.title || p.nav_title || `Session ${idx+1}`)}</option>
  `).join('');
}

function confirmAddGlobalExercise() {
  const athSelect = document.getElementById('pick-global-ath-select');
  const sessSelect = document.getElementById('pick-global-sess-select');
  if (!athSelect || !sessSelect || !pendingLibraryExercise) return;

  const athId = athSelect.value;
  const sessIdx = parseInt(sessSelect.value, 10);
  const athlete = studioAthletes.find(a => a.id === athId);

  if (!athlete || isNaN(sessIdx) || !athlete.pages[sessIdx]) {
    showToast(' Please select a valid athlete and workout session.');
    return;
  }

  const page = athlete.pages[sessIdx];
  if (!page.cards) page.cards = [];

  page.cards.push({
    category: pendingLibraryExercise.category || 'Primary Strength',
    exercise: pendingLibraryExercise.exercise,
    details: pendingLibraryExercise.details,
    weight: pendingLibraryExercise.weight || '',
    rest: pendingLibraryExercise.rest || '60s',
    rest_seconds: pendingLibraryExercise.rest_seconds || 60,
    total_sets: pendingLibraryExercise.total_sets || 3,
    cue: pendingLibraryExercise.cue || '',
    videos: pendingLibraryExercise.videos ? [...pendingLibraryExercise.videos] : []
  });

  saveAthletesToStorage();
  closeModal('pick-global-modal');
  closeModal('library-modal');
  showToast(` Added "${pendingLibraryExercise.exercise}" to ${athlete.name}'s ${page.title}!`);

  if (currentView === 'hub') renderHubView();
  else if (currentView === 'athlete') renderAthleteDetail();
}

// CREATE CUSTOM EXERCISE MODAL
function openCreateCustomExerciseModal() {
  const cCat = document.getElementById('custom-ex-category');
  const cGrp = document.getElementById('custom-ex-group');
  if (cCat) enhanceSelect(cCat);
  if (cGrp) enhanceSelect(cGrp);
  document.getElementById('custom-ex-name').value = '';
  document.getElementById('custom-ex-details').value = '3 x 10';
  document.getElementById('custom-ex-weight').value = 'Bodyweight';
  document.getElementById('custom-ex-rest').value = '60s';
  document.getElementById('custom-ex-cue').value = '';
  document.getElementById('custom-ex-video').value = '';
  openModal('create-custom-exercise-modal');
}

function saveCustomExercise() {
  const name = document.getElementById('custom-ex-name').value.trim();
  if (!name) {
    showToast('Please enter an exercise name.', MUI.warning);
    return;
  }

  const cat = document.getElementById('custom-ex-category').value;
  const group = document.getElementById('custom-ex-group').value;
  const details = document.getElementById('custom-ex-details').value.trim() || '3 x 10';
  const weight = document.getElementById('custom-ex-weight').value.trim() || 'Bodyweight';
  const rest = document.getElementById('custom-ex-rest').value.trim() || '60s';
  const cue = document.getElementById('custom-ex-cue').value.trim();
  const videoUrl = document.getElementById('custom-ex-video').value.trim();

  let restSec = 60;
  const matchSec = rest.match(/\\d+/);
  if (matchSec) restSec = parseInt(matchSec[0], 10);

  const newExercise = {
    group: group,
    category: cat,
    exercise: name,
    details: details,
    weight: weight,
    rest: rest,
    rest_seconds: restSec,
    total_sets: 3,
    cue: cue,
    videos: videoUrl ? [videoUrl] : [],
    isCustom: true
  };

  customLibrary.push(newExercise);
  try {
    localStorage.setItem('coach_studio_custom_library', JSON.stringify(customLibrary));
  } catch (e) {}

  closeModal('create-custom-exercise-modal');
  showToast(` Created and added "${name}" to Exercise Library!`);
  renderLibraryCards();
}

// PHONE PREVIEW MODAL & RESPONSIVE PROPORTIONAL SCALING
function fitPhonePreview() {
  const stage = document.querySelector('.preview-stage');
  const phone = document.getElementById('phone-chassis');
  if (!stage || !phone) return;
  const isWide = phone.classList.contains('wide');
  const phoneTotalH = isWide ? 932 + 24 : 844 + 24;
  const phoneTotalW = isWide ? 430 + 24 : 390 + 24;
  const availH = Math.max(320, (window.innerHeight || 800) - 130);
  const availW = Math.max(300, (window.innerWidth || 1200) - 40);
  const scale = Math.min(1, availH / phoneTotalH, availW / phoneTotalW);
  phone.style.transform = `translateX(-50%) scale(${scale})`;
  stage.style.width = `${Math.round(phoneTotalW * scale)}px`;
  stage.style.height = `${Math.round(phoneTotalH * scale)}px`;
}

function openPreviewModal() {
  const ath = getActiveAthlete();
  const iframe = document.getElementById('preview-iframe');
  if (iframe && ath) {
    iframe.srcdoc = generateClientHtml(ath);
  }
  openModal('preview-modal');
  setTimeout(fitPhonePreview, 60);
}

function setPreviewDevice(mode) {
  const phone = document.getElementById('phone-chassis');
  const btnStd = document.getElementById('btn-device-std');
  const btnMax = document.getElementById('btn-device-max');
  if (!phone) return;
  if (mode === 'wide') {
    phone.classList.add('wide');
    if (btnMax) btnMax.classList.add('btn-header-primary');
    if (btnStd) btnStd.classList.remove('btn-header-primary');
  } else {
    phone.classList.remove('wide');
    if (btnStd) btnStd.classList.add('btn-header-primary');
    if (btnMax) btnMax.classList.remove('btn-header-primary');
  }
  fitPhonePreview();
}

if (typeof window !== 'undefined' && window.addEventListener) {
  window.addEventListener('resize', () => {
    const modal = document.getElementById('preview-modal');
    if (modal && modal.classList.contains('show')) {
      fitPhonePreview();
    }
  });
}

// EXPORT CLIENT HTML APP CONTROLLERS
let exportSelectedAthleteId = null;
let exportSelectedSessionIds = new Set();

function openExportSessionsModal(athleteId = null) {
  const targetId = athleteId || activeAthleteId;
  const ath = studioAthletes.find(a => a.id === targetId) || getActiveAthlete();
  if (!ath) {
    showToast('Please select an athlete to export.');
    return;
  }
  exportSelectedAthleteId = ath.id;

  const athleteNameEl = document.getElementById('export-modal-athlete-name');
  if (athleteNameEl) athleteNameEl.textContent = ath.name;

  const filenameInput = document.getElementById('export-filename-input');
  if (filenameInput) {
    filenameInput.value = `${ath.name.replace(/\\s+/g, '_')}_Program.html`;
  }

  const listContainer = document.getElementById('export-sessions-list');
  if (!listContainer) return;

  const pages = ath.pages || [];
  exportSelectedSessionIds.clear();

  if (pages.length === 0) {
    listContainer.innerHTML = `
      <div style="padding:24px; text-align:center; color:var(--text-muted); background:rgba(255,255,255,0.02); border-radius:10px; border:1px dashed rgba(255,255,255,0.1);">
        No sessions found for this athlete. Add a session first before exporting.
      </div>`;
    updateExportButtonState();
    openModal('export-sessions-modal');
    return;
  }

  pages.forEach((p, idx) => {
    const pageId = p.id || `page-${idx}`;
    p.id = pageId;
    exportSelectedSessionIds.add(pageId);
  });

  listContainer.innerHTML = pages.map((p, idx) => {
    const pageId = p.id || `page-${idx}`;
    const title = p.nav_title || p.title || `Session ${idx + 1}`;
    const fullTitle = p.title && p.title !== p.nav_title ? `${p.title}` : title;
    const cardCount = (p.cards || []).length;
    let setCount = 0;
    if (p.cards) {
      p.cards.forEach(c => { setCount += (c.total_sets || 3); });
    }
    const isChecked = exportSelectedSessionIds.has(pageId);

    return `
      <label class="export-session-item ${isChecked ? 'selected' : ''}" id="export-item-${pageId}" style="display:flex; align-items:center; justify-content:space-between; padding:10px 14px; background:rgba(255,255,255,0.03); border:1px solid rgba(255,255,255,0.08); border-radius:10px; cursor:pointer; transition:all 0.18s ease;">
        <div style="display:flex; align-items:center; gap:12px; flex:1; min-width:0;">
          <input type="checkbox" class="export-session-checkbox" data-page-id="${pageId}" ${isChecked ? 'checked' : ''} onchange="toggleExportSession('${pageId}', this.checked)" style="width:18px; height:18px; accent-color:var(--teal); cursor:pointer;">
          <div style="overflow:hidden; text-overflow:ellipsis; white-space:nowrap;">
            <div style="font-weight:700; font-size:0.92em; color:var(--text);">${escapeHtml(fullTitle)}</div>
            <div style="font-size:0.78em; color:var(--text-muted); margin-top:2px;">
              ${p.type === 'table' ? 'Overview / Table View' : `${cardCount} exercises  -  ${setCount} sets`}
            </div>
          </div>
        </div>
        <span class="badge" style="font-size:0.75em; padding:3px 8px; border-radius:6px; background:rgba(0,229,255,0.1); color:var(--teal); border:1px solid rgba(0,229,255,0.25);">
          #${idx + 1}
        </span>
      </label>
    `;
  }).join('');

  updateExportButtonState();
  openModal('export-sessions-modal');
}

function toggleExportSession(pageId, isChecked) {
  if (isChecked) {
    exportSelectedSessionIds.add(pageId);
  } else {
    exportSelectedSessionIds.delete(pageId);
  }
  const itemEl = document.getElementById(`export-item-${pageId}`);
  if (itemEl) {
    if (isChecked) itemEl.classList.add('selected');
    else itemEl.classList.remove('selected');
  }
  updateExportButtonState();
}

function selectAllExportSessions(select) {
  const checkboxes = document.querySelectorAll('.export-session-checkbox');
  checkboxes.forEach(cb => {
    cb.checked = select;
    const pageId = cb.getAttribute('data-page-id');
    if (pageId) {
      if (select) exportSelectedSessionIds.add(pageId);
      else exportSelectedSessionIds.delete(pageId);
      const itemEl = document.getElementById(`export-item-${pageId}`);
      if (itemEl) {
        if (select) itemEl.classList.add('selected');
        else itemEl.classList.remove('selected');
      }
    }
  });
  updateExportButtonState();
}

function updateExportButtonState() {
  const btn = document.getElementById('export-confirm-btn');
  if (!btn) return;
  const count = exportSelectedSessionIds.size;
  if (count === 0) {
    btn.disabled = true;
    btn.style.opacity = '0.5';
    btn.style.pointerEvents = 'none';
    btn.innerHTML = `${MUI.export} Select at least 1 session`;
  } else {
    btn.disabled = false;
    btn.style.opacity = '1';
    btn.style.pointerEvents = 'auto';
    btn.innerHTML = `${MUI.export} Download ${count} Session${count === 1 ? '' : 's'} (.html)`;
  }
}

async function confirmExportSelectedSessions(exportMode = 'download') {
  const ath = studioAthletes.find(a => a.id === exportSelectedAthleteId) || getActiveAthlete();
  if (!ath) return;

  const selectedIds = Array.from(exportSelectedSessionIds);
  if (selectedIds.length === 0) {
    showToast('Please select at least one session to export.');
    return;
  }

  let filename = (document.getElementById('export-filename-input')?.value || '').trim();
  if (!filename) {
    filename = `${ath.name.replace(/\\s+/g, '_')}_Program.html`;
  }
  if (!filename.toLowerCase().endsWith('.html')) {
    filename += '.html';
  }

  const html = generateClientHtml(ath, selectedIds);
  const blob = new Blob([html], { type: 'text/html;charset=utf-8' });

  if (exportMode === 'share') {
    const file = new File([blob], filename, { type: 'text/html' });
    if (navigator.canShare && navigator.canShare({ files: [file] })) {
      try {
        await navigator.share({
          files: [file],
          title: `${ath.name} - Workout Program`,
          text: `Here is your workout program for ${ath.name}. Open this file in your browser to start!`
        });
        closeModal('export-sessions-modal');
        showToast('Shared successfully via WhatsApp / native share!');
        return;
      } catch (err) {
        if (err.name !== 'AbortError') {
          console.warn('Share error:', err);
        } else {
          return; // User cancelled share sheet
        }
      }
    } else if (navigator.share) {
      try {
        // Fallback for browsers supporting link share
        const a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        closeModal('export-sessions-modal');
        showToast(`Saved ${filename}. Attach it in WhatsApp!`);
        return;
      } catch (e) {}
    }
  }

  // Standard download mode or fallback
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);

  closeModal('export-sessions-modal');
  showToast(`Exported ${selectedIds.length} session${selectedIds.length === 1 ? '' : 's'} to ${filename}!`);
}

function exportAthleteClientHtml(athleteId = null) {
  openExportSessionsModal(athleteId);
}

function quickShareAthleteWhatsApp(athleteId = null) {
  const targetId = athleteId || activeAthleteId;
  const ath = studioAthletes.find(a => a.id === targetId);
  if (!ath) return;
  openExportSessionsModal(targetId);
  const shareBtn = document.getElementById('export-share-btn');
  if (shareBtn) {
    shareBtn.focus();
  }
}

// CSV TEMPLATE DOWNLOADER & BATCH IMPORT CONTROLLER
function downloadCsvTemplate() {
  const headers = ['Exercise Name', 'Category', 'Sets x Reps', 'Weight', 'Rest', 'Coaching Cue', 'Superset'];
  const sampleRows = [
    ['Dumbbell Bench Press', 'Primary Strength', '4 x 8-10', '16kg', '90s', 'Pin shoulder blades back and drive through feet.', 'A1'],
    ['Lat Pulldown', 'Primary Strength', '3 x 10-12', '45kg', '75s', 'Drive elbows down to hips with controlled eccentric.', 'A2'],
    ['Barbell Back Squat', 'Primary Strength', '4 x 6-8', '80kg', '120s', 'Keep chest tall and track knees over toes.', ''],
    ['Romanian Deadlift', 'Secondary Strength', '3 x 8-10', '60kg', '90s', 'Hinge deep at hips with neutral spine.', ''],
    ['Bent-Elbow Band Pull-Aparts', 'Warm-Up', '2 x 15', 'Light Band', '30s', 'Squeeze upper back without shrugging.', ''],
    ['Lateral Raises', 'Accessory', '3 x 12-15', '6kg', '60s', 'Lead with elbows and pause at top.', '']
  ];

  function escapeCsvValue(val) {
    if (val === null || val === undefined) return '';
    const str = String(val);
    if (str.includes(',') || str.includes('"') || str.includes('\\n') || str.includes('\\r')) {
      return `"${str.replace(/"/g, '""')}"`;
    }
    return str;
  }

  const csvLines = [
    headers.map(escapeCsvValue).join(','),
    ...sampleRows.map(row => row.map(escapeCsvValue).join(','))
  ];

  const csvContent = '\\uFEFF' + csvLines.join('\\r\\n');
  const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'workout_exercise_template.csv';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
  showToast('Workout CSV template downloaded!');
}

function openBatchImportModal() {
  const textarea = document.getElementById('batch-import-text');
  if (textarea) textarea.value = '';
  previewBatchImport();
  openModal('batch-import-modal');
}

function parseBatchImportText(text) {
  if (!text || !text.trim()) return [];
  const lines = text.trim().split(/\\r?\\n/);
  const exercises = [];

  for (let i = 0; i < lines.length; i++) {
    const rawLine = lines[i].trim();
    if (!rawLine) continue;

    let parts = [];
    if (rawLine.includes('\\t')) {
      parts = rawLine.split('\\t').map(s => s.trim());
    } else {
      const matches = rawLine.match(/(".*?"|[^",]+)(?=\\s*,|\\s*$)/g);
      parts = matches ? matches.map(s => s.replace(/^"|"$/g, '').trim()) : rawLine.split(',').map(s => s.trim());
    }

    if (parts.length === 0 || !parts[0]) continue;

    const firstCol = parts[0].toLowerCase();
    if (i === 0 && (firstCol === 'exercise name' || firstCol === 'exercise' || firstCol === 'name')) {
      continue;
    }

    let name = parts[0];
    let category = 'Primary Strength';
    let details = '3 x 10';
    let weight = '';
    let rest = '60s';
    let cue = '';
    let superset = '';

    if (parts.length >= 7) {
      category = parts[1] || category;
      details = parts[2] || details;
      weight = parts[3] || '';
      rest = parts[4] || rest;
      cue = parts[5] || '';
      superset = parts[6] || '';
    } else if (parts.length === 6) {
      details = parts[1] || details;
      weight = parts[2] || '';
      rest = parts[3] || rest;
      cue = parts[4] || '';
      superset = parts[5] || '';
    } else if (parts.length === 5) {
      details = parts[1] || details;
      weight = parts[2] || '';
      rest = parts[3] || rest;
      cue = parts[4] || '';
    } else if (parts.length >= 2) {
      details = parts[1] || details;
      if (parts[2]) weight = parts[2];
      if (parts[3]) rest = parts[3];
      if (parts[4]) cue = parts[4];
    }

    let totalSets = 3;
    const setsMatch = details.match(/^(\\d+)/);
    if (setsMatch) totalSets = parseInt(setsMatch[1], 10) || 3;

    let restSec = 60;
    const restMatch = rest.match(/(\\d+)/);
    if (restMatch) {
      restSec = parseInt(restMatch[1], 10);
      if (rest.toLowerCase().includes('m')) restSec *= 60;
    }

    exercises.push({
      category: category,
      exercise: name,
      details: details,
      weight: weight,
      rest: rest,
      rest_seconds: restSec,
      total_sets: totalSets,
      cue: cue,
      superset: superset,
      videos: []
    });
  }
  return exercises;
}

function previewBatchImport() {
  const textarea = document.getElementById('batch-import-text');
  const countEl = document.getElementById('batch-import-preview-count');
  const listEl = document.getElementById('batch-import-preview-list');
  if (!textarea || !countEl || !listEl) return;

  const exercises = parseBatchImportText(textarea.value);
  countEl.textContent = `${exercises.length} exercise${exercises.length === 1 ? '' : 's'} detected`;

  if (exercises.length === 0) {
    listEl.innerHTML = '<span style="color:var(--text-dim);">Paste text or CSV data above to see preview.</span>';
    return;
  }

  listEl.innerHTML = exercises.map((ex, idx) => `
    <div style="padding:4px 0; border-bottom:1px solid rgba(255,255,255,0.05); display:flex; justify-content:space-between; align-items:center;">
      <span><strong>#${idx + 1} ${escapeHtml(ex.exercise)}</strong> <span style="color:var(--text-muted); font-size:0.9em;">(${escapeHtml(ex.details)}${ex.weight ? '  -  ' + escapeHtml(ex.weight) : ''}${ex.rest ? '  -  ' + escapeHtml(ex.rest) : ''})</span></span>
      <span style="font-size:0.8em; color:var(--teal);">${escapeHtml(ex.category)}</span>
    </div>
  `).join('');
}

function confirmBatchImport() {
  const textarea = document.getElementById('batch-import-text');
  if (!textarea) return;
  const exercises = parseBatchImportText(textarea.value);

  if (exercises.length === 0) {
    showToast('No valid exercises detected to insert.');
    return;
  }

  const ath = getActiveAthlete();
  if (!ath || !ath.pages || !ath.pages[activeSessionIndex]) {
    showToast('Please select a valid session first.');
    return;
  }

  const page = ath.pages[activeSessionIndex];
  if (!page.cards) page.cards = [];
  page.cards.push(...exercises);

  saveAthletesToStorage();
  renderBuilderView();
  closeModal('batch-import-modal');
  showToast(`Successfully inserted ${exercises.length} exercise${exercises.length === 1 ? '' : 's'}!`);
}

function openBackupModal() {
  openModal('backup-modal');
}

function exportStudioBackupJson() {
  const backupData = {
    version: '2.0',
    exportDate: new Date().toISOString(),
    athletes: studioAthletes,
    sessionTemplates: sessionTemplates,
    customLibrary: customLibrary
  };
  const blob = new Blob([JSON.stringify(backupData, null, 2)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `Coach_Studio_Backup_${new Date().toISOString().slice(0, 10)}.json`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
  showToast('Studio backup (.json) downloaded!');
}

function handleRestoreFileSelected(event) {
  const file = event.target.files && event.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = function(e) {
    try {
      const data = JSON.parse(e.target.result);
      if (data && Array.isArray(data.athletes)) {
        studioAthletes = data.athletes;
        if (Array.isArray(data.sessionTemplates)) sessionTemplates = data.sessionTemplates;
        if (Array.isArray(data.customLibrary)) {
          customLibrary = data.customLibrary;
          try {
            localStorage.setItem('coach_studio_custom_library', JSON.stringify(customLibrary));
          } catch (err) {}
        }
        saveAthletesToStorage();
        saveTemplatesToStorage();
        renderHubView();
        closeModal('backup-modal');
        showToast('Successfully restored Coach Studio data from backup!');
      } else {
        showToast('Invalid backup file format.');
      }
    } catch (err) {
      showToast('Failed to read backup file.');
    }
  };
  reader.readAsText(file);
}

function generateClientHtml(athlete, selectedPageIds = null) {
  const clientName = athlete.name || 'Athlete';
  const storageKey = 'fitness_tracker_' + clientName.toLowerCase().replace(/[^a-z0-9]/g, '_');
  let totalProgramSets = 0;
  let globalSetIdx = 0;

  const allPages = athlete.pages || [];
  const pages = (selectedPageIds && Array.isArray(selectedPageIds) && selectedPageIds.length > 0)
    ? allPages.filter((p, idx) => {
        const pid = p.id || `page-${idx}`;
        return selectedPageIds.includes(pid);
      })
    : allPages;
  pages.forEach(p => {
    if (p.cards) {
      p.cards.forEach(c => { totalProgramSets += (c.total_sets || 3); });
    }
  });

  const navHtml = pages.map((p, idx) => {
    const active = idx === 0 ? 'active' : '';
    const title = escapeHtml(p.nav_title || p.title || `Session ${idx+1}`);
    const pageId = p.id || `page-${idx}`;
    return `    <button class="nav-btn ${active}" data-page="${pageId}" onclick="showPage('${pageId}', this)">${title}</button>`;
  }).join('\\n');

  const pagesHtml = pages.map((p, idx) => {
    const active = idx === 0 ? 'active' : '';
    const pageId = p.id || `page-${idx}`;
    let pageContent = `<div id="${pageId}" class="session-page ${active}" data-type="${p.type || 'exercises'}">\\n`;
    pageContent += `  <div class="session-title">${escapeHtml(p.title || p.nav_title || 'Session')}</div>\\n`;

    if (p.type === 'table') {
      pageContent += `  <div class="plan-container">\\n`;
      (p.cards || []).forEach(c => {
        pageContent += `    <div class="day-card">\\n`;
        pageContent += `      <div class="day-name">${escapeHtml(c.category || 'Day')}</div>\\n`;
        if (c.exercise) pageContent += `      <div class="day-focus">${escapeHtml(c.exercise)}</div>\\n`;
        if (c.details) pageContent += `      <div class="day-details">${escapeHtml(c.details)}</div>\\n`;
        pageContent += `    </div>\\n`;
      });
      pageContent += `  </div>\\n`;
    } else {
      const sessionSets = (p.cards || []).reduce((acc, c) => acc + (c.total_sets || 3), 0);
      pageContent += `  <div class="session-info-bar"><span class="session-badge">${(p.cards || []).length} Exercises</span><span class="session-badge">${sessionSets} Sets</span></div>\\n`;

      (p.cards || []).forEach(c => {
        const totalSets = c.total_sets || 3;
        pageContent += `  <div class="exercise-card">\\n`;
        if (c.category) pageContent += `    <div class="category">${escapeHtml(c.category)}</div>\\n`;
        pageContent += `    <div class="title">${escapeHtml(c.exercise)}</div>\\n`;
        if (c.details) pageContent += `    <div class="details">${escapeHtml(c.details)}</div>\\n`;

        if (c.weight || c.rest) {
          pageContent += `    <div class="badge-container">\\n`;
          if (c.weight) pageContent += `      <div class="badge badge-weight">${escapeHtml(c.weight)}</div>\\n`;
          if (c.rest) {
            const sec = c.rest_seconds || 60;
            pageContent += `      <button class="badge badge-rest" onclick="startRestFromBadge(${sec}, this)">${escapeHtml(c.rest)} <span class="badge-tap-hint">▶ Start</span></button>\\n`;
          }
          pageContent += `    </div>\\n`;
        }

        if (c.cue) {
          pageContent += `    <div class="cue"><span class="cue-icon"></span> ${escapeHtml(c.cue)}</div>\\n`;
        }

        if (c.videos && c.videos.length > 0) {
          pageContent += `    <div class="links">\\n`;
          c.videos.forEach((v, vIdx) => {
            const safeTitle = c.exercise.replace(/'/g, "\\\\'");
            pageContent += `      <button class="video-btn" onclick="openVideoModal('${escapeHtml(v)}', '${escapeHtml(safeTitle)}')">▶ Watch Demo ${vIdx+1}</button>\\n`;
          });
          pageContent += `    </div>\\n`;
        }

        pageContent += `    <div class="sets-container">\\n`;
        pageContent += `      <div class="sets-header"><span class="sets-title">Track Sets</span><span class="sets-subtitle">Log Weight & Reps</span></div>\\n`;
        pageContent += `      <div class="sets-grid">\\n`;
        for (let s = 1; s <= totalSets; s++) {
          pageContent += `        <div class="set-row">\\n`;
          pageContent += `          <label class="set-checkbox" title="Check Set ${s}"><input type="checkbox" data-set-id="${globalSetIdx}"><span>Set ${s}</span></label>\\n`;
          pageContent += `          <div class="set-inputs">\\n`;
          pageContent += `            <input type="text" class="set-input set-weight" data-set-id="${globalSetIdx}" placeholder="kg" inputmode="decimal">\\n`;
          pageContent += `            <input type="text" class="set-input set-reps" data-set-id="${globalSetIdx}" placeholder="reps" inputmode="numeric">\\n`;
          pageContent += `          </div>\\n`;
          pageContent += `        </div>\\n`;
          globalSetIdx++;
        }
        pageContent += `      </div>\\n    </div>\\n  </div>\\n`;
      });

      pageContent += `  <div class="session-notes-card">\\n`;
      pageContent += `    <div class="notes-header">Session Notes for ${escapeHtml(p.nav_title || 'Session')}</div>\\n`;
      pageContent += `    <textarea class="session-notes-input" data-session-id="${pageId}" placeholder="Log personal reflections, weights used, or fatigue notes for next week..." oninput="saveSessionNote('${pageId}', this.value)"></textarea>\\n`;
      pageContent += `  </div>\\n`;
      pageContent += `  <button class="btn-finish-workout" onclick="finishSession()" style="margin-top:16px;">Finish Session</button>\\n`;
    }

    pageContent += `</div>`;
    return pageContent;
  }).join('\\n');

  let jsCode = CLIENT_JS_TEMPLATE
    .replace('{storage_key}', storageKey)
    .replace('{total_program_sets}', totalProgramSets);

  const clientExtraScript = `
// -------------------------------------------------------------
// REST TIMER VISIBILITY & PERSISTENCE CONTROLLER
// -------------------------------------------------------------
function toggleTimerVisibility(forcedState) {
  var timerBar = document.getElementById('timer-bar');
  if (!timerBar) return;

  var isCurrentlyHidden = timerBar.classList.contains('hidden');
  var shouldShow = (typeof forcedState === 'boolean') ? forcedState : isCurrentlyHidden;

  var menuTimerText = document.getElementById('menu-timer-text');
  var btn = document.getElementById('btn-toggle-timer');
  var btnText = document.getElementById('timer-btn-text');

  if (shouldShow) {
    timerBar.classList.remove('hidden');
    if (window.safeStorage) safeStorage.setItem(STORAGE_KEY + '_show_timer', 'true');
    if (menuTimerText) menuTimerText.textContent = 'Rest Timer (On)';
    if (btn) {
      btn.classList.add('btn-active-timer');
      btn.classList.remove('btn-inactive-timer');
      btn.title = 'Hide Rest Timer';
    }
    if (btnText) btnText.textContent = 'Timer';
    if (typeof forcedState !== 'boolean') {
      showToast('Rest timer enabled');
    }
  } else {
    timerBar.classList.add('hidden');
    if (window.safeStorage) safeStorage.setItem(STORAGE_KEY + '_show_timer', 'false');
    if (menuTimerText) menuTimerText.textContent = 'Rest Timer (Off)';
    if (btn) {
      btn.classList.remove('btn-active-timer');
      btn.classList.add('btn-inactive-timer');
      btn.title = 'Show Rest Timer';
    }
    if (btnText) btnText.textContent = 'Timer Off';
    showToast('Rest timer hidden');
  }
}

function initTimerVisibility() {
  var saved = window.safeStorage ? safeStorage.getItem(STORAGE_KEY + '_show_timer') : null;
  if (saved === 'false') {
    toggleTimerVisibility(false);
  } else {
    var menuTimerText = document.getElementById('menu-timer-text');
    if (menuTimerText) menuTimerText.textContent = 'Rest Timer (On)';
    var btn = document.getElementById('btn-toggle-timer');
    if (btn) {
      btn.classList.add('btn-active-timer');
      btn.classList.remove('btn-inactive-timer');
      btn.title = 'Hide Rest Timer';
    }
  }
  initTimerBadgeObserver();
}

function initTimerBadgeObserver() {
  var startBtnEl = document.getElementById('start-btn');
  if (startBtnEl && window.MutationObserver) {
    var obs = new MutationObserver(function() {
      var isTimerRunning = startBtnEl.classList.contains('running');
      var badge = document.getElementById('session-timer-badge');
      if (badge) {
        if (isTimerRunning) badge.classList.add('active');
        else badge.classList.remove('active');
      }
    });
    obs.observe(startBtnEl, { attributes: true, attributeFilter: ['class'] });
  }
}

function toggleSessionMenu(event) {
  if (event) {
    event.stopPropagation();
    event.preventDefault();
  }
  var menu = document.getElementById('session-dropdown-menu');
  var btn = document.getElementById('btn-session-menu');
  if (!menu) return;
  var isShowing = menu.classList.contains('show');
  if (isShowing) {
    closeSessionMenu();
  } else {
    menu.classList.add('show');
    if (btn) btn.classList.add('active');
  }
}

function closeSessionMenu() {
  var menu = document.getElementById('session-dropdown-menu');
  var btn = document.getElementById('btn-session-menu');
  if (menu) menu.classList.remove('show');
  if (btn) btn.classList.remove('active');
}

document.addEventListener('click', function(e) {
  var wrap = document.querySelector('.session-menu-wrap');
  if (wrap && !wrap.contains(e.target)) {
    closeSessionMenu();
  }
});

document.addEventListener('keydown', function(e) {
  if (e.key === 'Escape') {
    closeSessionMenu();
  }
});

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', initTimerVisibility);
} else {
  initTimerVisibility();
}

// Automatically reveal timer if athlete clicks any rest badge
var origStartRestFromBadge = window.startRestFromBadge;
window.startRestFromBadge = function(seconds, el) {
  var timerBar = document.getElementById('timer-bar');
  if (timerBar && timerBar.classList.contains('hidden')) {
    toggleTimerVisibility(true);
  }
  if (typeof origStartRestFromBadge === 'function') {
    origStartRestFromBadge(seconds, el);
  }
};

// -------------------------------------------------------------
// CLIENT WORKOUT UTILITIES (Recap, Plates, Focus, History)
// -------------------------------------------------------------
function closeClientModal(id) {
  var el = document.getElementById(id);
  if (el) {
    el.classList.remove('show');
    el.classList.remove('active');
  }
}

function openClientModal(id) {
  var el = document.getElementById(id);
  if (el) {
    el.classList.add('show');
    el.classList.add('active');
  }
}

function openRecapModal() {
  var activePage = document.querySelector('.session-page.active') || document.querySelector('.session-page');
  var titleEl = activePage ? activePage.querySelector('.session-title') : null;
  var sessionTitle = titleEl ? titleEl.textContent : 'Workout';
  var cards = activePage ? activePage.querySelectorAll('.exercise-card') : [];
  var h1El = document.querySelector('h1');
  var athleteHeading = h1El ? h1El.textContent.trim() : 'Athlete';

  var lines = [];
  lines.push('Workout Recap: ' + athleteHeading);
  lines.push(sessionTitle + ' (' + new Date().toISOString().slice(0, 10) + ')');
  lines.push('------------------------');

  cards.forEach(function(card) {
    var tEl = card.querySelector('.title');
    var title = tEl ? tEl.textContent.trim() : 'Exercise';
    var rows = card.querySelectorAll('.set-row');
    var loggedSets = [];

    rows.forEach(function(row) {
      var cb = row.querySelector('input[type=checkbox]');
      if (cb && cb.checked) {
        var wEl = row.querySelector('.set-weight');
        var rEl = row.querySelector('.set-reps');
        var w = wEl ? wEl.value.trim() : '';
        var r = rEl ? rEl.value.trim() : '';
        if (w || r) loggedSets.push((w ? w + 'kg' : '') + (r ? ' x ' + r : ''));
        else loggedSets.push('Done');
      }
    });

    if (loggedSets.length > 0) {
      lines.push(title + ': ' + loggedSets.length + ' sets (' + loggedSets.join(', ') + ')');
    }
  });

  var notesEl = activePage ? activePage.querySelector('.session-notes-input') : null;
  var notes = notesEl ? notesEl.value.trim() : '';
  if (notes) {
    lines.push('');
    lines.push('Coach Notes: ' + notes);
  }

  var summary = lines.join(String.fromCharCode(10));
  var recapTextEl = document.getElementById('recap-text');
  if (recapTextEl) recapTextEl.value = summary;
  openClientModal('recap-modal');
}

function copyRecapToClipboard() {
  var recapTextEl = document.getElementById('recap-text');
  var text = recapTextEl ? recapTextEl.value : '';
  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(text).then(function() {
      alert('Workout summary copied to clipboard! Paste directly into WhatsApp or SMS.');
    });
  } else {
    alert('Workout summary: ' + text);
  }
}

function openWhatsAppChat() {
  var recapTextEl = document.getElementById('recap-text');
  var text = recapTextEl ? recapTextEl.value : '';
  window.open('https://api.whatsapp.com/send?text=' + encodeURIComponent(text), '_blank');
}


function togglePlateBarDropdown(e) {
  if (e) e.stopPropagation();
  var trigger = document.getElementById('plate-bar-trigger');
  var menu = document.getElementById('plate-bar-menu');
  if (!trigger || !menu) return;
  var isOpen = menu.classList.contains('show');
  if (isOpen) {
    menu.classList.remove('show');
    trigger.classList.remove('active');
  } else {
    menu.classList.add('show');
    trigger.classList.add('active');
  }
}

function selectPlateBar(val, label) {
  var hidden = document.getElementById('plate-bar-weight');
  var triggerText = document.getElementById('plate-bar-trigger-text');
  var trigger = document.getElementById('plate-bar-trigger');
  var menu = document.getElementById('plate-bar-menu');

  if (hidden) hidden.value = val;
  if (triggerText) triggerText.textContent = label;

  var opt20 = document.getElementById('bar-opt-20');
  var opt15 = document.getElementById('bar-opt-15');
  var chk20 = document.getElementById('bar-chk-20');
  var chk15 = document.getElementById('bar-chk-15');

  if (val == 20) {
    if (opt20) opt20.classList.add('active');
    if (opt15) opt15.classList.remove('active');
    if (chk20) chk20.style.display = 'block';
    if (chk15) chk15.style.display = 'none';
  } else {
    if (opt15) opt15.classList.add('active');
    if (opt20) opt20.classList.remove('active');
    if (chk15) chk15.style.display = 'block';
    if (chk20) chk20.style.display = 'none';
  }

  if (menu) menu.classList.remove('show');
  if (trigger) trigger.classList.remove('active');
  calcPlates();
}

document.addEventListener('click', function(e) {
  var menu = document.getElementById('plate-bar-menu');
  var trigger = document.getElementById('plate-bar-trigger');
  if (menu && trigger && !trigger.contains(e.target) && !menu.contains(e.target)) {
    menu.classList.remove('show');
    trigger.classList.remove('active');
  }
});

function openPlateModal() {
  calcPlates();
  openClientModal('plate-modal');
}

function calcPlates() {
  var targetEl = document.getElementById('plate-target-weight');
  var barEl = document.getElementById('plate-bar-weight');
  var target = targetEl ? (parseFloat(targetEl.value) || 20) : 20;
  var bar = barEl ? (parseFloat(barEl.value) || 20) : 20;
  var chipsContainer = document.getElementById('plate-result-chips');
  var sideWeightEl = document.getElementById('plate-side-weight');
  var warmupList = document.getElementById('plate-warmup-list');

  if (target < bar) target = bar;
  var perSide = (target - bar) / 2;
  if (sideWeightEl) sideWeightEl.textContent = 'Each side loads: ' + perSide.toFixed(1) + ' kg';

  var available = [20, 15, 10, 5, 2.5, 1.25];
  var count = {};
  var rem = perSide;

  available.forEach(function(p) {
    var c = Math.floor(rem / p);
    if (c > 0) {
      count[p] = c;
      rem = +(rem - (c * p)).toFixed(2);
    }
  });

  var chipsHtml = '';
  available.forEach(function(p) {
    if (count[p]) {
      var cls = 'plate-' + String(p).replace('.', '_');
      chipsHtml += '<span class="plate-chip ' + cls + '">' + count[p] + 'x ' + p + 'kg</span>';
    }
  });
  if (chipsContainer) {
    chipsContainer.innerHTML = chipsHtml || '<span style="color:#94A3B8; font-size:0.8em;">Bar Only (No plates needed)</span>';
  }

  if (warmupList) {
    warmupList.innerHTML = [
      ' -  Set 1 (Empty Bar): ' + bar + 'kg x 10',
      ' -  Set 2 (50%): ' + (Math.round((target * 0.5) / 2.5) * 2.5) + 'kg x 5',
      ' -  Set 3 (70%): ' + (Math.round((target * 0.7) / 2.5) * 2.5) + 'kg x 3',
      ' -  Set 4 (85%): ' + (Math.round((target * 0.85) / 2.5) * 2.5) + 'kg x 1',
      ' -  Working Sets (100%): ' + target + 'kg'
    ].join('<br>');
  }
}

function finishAndLogWorkout(pageId) {
  var activePage = document.getElementById(pageId) || document.querySelector('.session-page.active') || document.querySelector('.session-page');
  var titleEl = activePage ? activePage.querySelector('.session-title') : null;
  var title = titleEl ? titleEl.textContent : 'Workout Session';
  var checkboxes = activePage ? activePage.querySelectorAll('input[type=checkbox]') : [];
  var checkedCount = 0;
  checkboxes.forEach(function(cb) { if (cb.checked) checkedCount++; });

  var entry = {
    date: new Date().toLocaleDateString(),
    time: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }),
    sessionTitle: title,
    setsCompleted: checkedCount,
    totalSets: checkboxes.length
  };

  try {
    var history = JSON.parse(safeStorage.getItem('client_workout_history') || '[]');
    history.unshift(entry);
    safeStorage.setItem('client_workout_history', JSON.stringify(history.slice(0, 30)));
  } catch (e) {}

  if (window.fireCelebrationConfetti) window.fireCelebrationConfetti();
  alert('Awesome job! Workout completed and saved to your history.');
}

function openHistoryModal() {
  var list = document.getElementById('history-log-list');
  var history = [];
  try {
    history = JSON.parse(safeStorage.getItem('client_workout_history') || '[]');
  } catch (e) {}

  if (list) {
    if (history.length === 0) {
      list.innerHTML = '<div style="color:#94A3B8; font-size:0.85em; text-align:center; padding:20px;">No workout sessions logged yet. Complete a workout and tap "Finish & Log Workout" to save.</div>';
    } else {
      list.innerHTML = history.map(function(h) {
        return '<div style="background:#070A10; border:1px solid rgba(255,255,255,0.08); border-radius:10px; padding:10px 14px;">' +
          '<div style="display:flex; justify-content:space-between; font-weight:800; color:#F8FAFC; font-size:0.9em;">' +
            '<span>' + h.sessionTitle + '</span>' +
            '<span style="color:#00E5FF; font-size:0.82em;">' + h.date + '</span>' +
          '</div>' +
          '<div style="font-size:0.75em; color:#94A3B8; margin-top:2px;">' + h.setsCompleted + ' of ' + h.totalSets + ' sets logged (' + (h.time || '') + ')</div>' +
        '</div>';
      }).join('');
    }
  }
  openClientModal('history-modal');
}

// -------------------------------------------------------------
// FULL-SCREEN FOCUS WORKOUT COCKPIT CONTROLLER (UPGRADED)
// -------------------------------------------------------------
var focusIndex = 0;
var focusCards = [];
var currentFocusRestSeconds = 60;

function escapeFocusHtml(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function openFocusMode() {
  var activePage = document.querySelector('.session-page.active') || document.querySelector('.session-page');
  if (!activePage) return;
  focusCards = Array.from(activePage.querySelectorAll('.exercise-card'));
  if (focusCards.length === 0) {
    if (typeof showToast === 'function') showToast('No exercises in this session for Focus Mode.');
    else alert('No exercises in this session for Focus Mode.');
    return;
  }
  // Find first exercise with uncompleted sets
  var targetIdx = 0;
  for (var i = 0; i < focusCards.length; i++) {
    var cbs = focusCards[i].querySelectorAll('.set-checkbox input[type="checkbox"]');
    var allChecked = cbs.length > 0 && Array.from(cbs).every(function(c) { return c.checked; });
    if (!allChecked) {
      targetIdx = i;
      break;
    }
  }
  focusIndex = targetIdx;
  renderFocusExercise();
  var overlay = document.getElementById('focus-overlay');
  if (overlay) {
    overlay.classList.add('show');
    overlay.classList.add('active');
    document.body.classList.add('focus-open');
  }
}

function closeFocusMode() {
  var overlay = document.getElementById('focus-overlay');
  if (overlay) {
    overlay.classList.remove('show');
    overlay.classList.remove('active');
    document.body.classList.remove('focus-open');
  }
}

function goToFocusExercise(idx) {
  if (idx < 0 || idx >= focusCards.length) return;
  focusIndex = idx;
  renderFocusExercise();
  var sb = document.getElementById('focus-scroll-body');
  if (sb) sb.scrollTop = 0;
}

function renderFocusExercise() {
  var card = focusCards[focusIndex];
  if (!card) return;

  var bText = document.getElementById('focus-badge-text');
  var catEl = document.getElementById('focus-ex-category');
  var tEl = document.getElementById('focus-ex-title');
  var detEl = document.getElementById('focus-ex-details');
  var cueCard = document.getElementById('focus-cue-card');
  var cueText = document.getElementById('focus-cue-text');
  var videoArea = document.getElementById('focus-video-area');

  var cardCat = card.querySelector('.category');
  var cardTitle = card.querySelector('.title');
  var cardDet = card.querySelector('.details');
  var origCue = card.querySelector('.cue');

  if (bText) bText.textContent = 'EXERCISE ' + (focusIndex + 1) + ' OF ' + focusCards.length;
  if (catEl) catEl.textContent = cardCat ? cardCat.textContent.trim() : 'EXERCISE';
  if (tEl) tEl.textContent = cardTitle ? cardTitle.textContent.trim() : 'Exercise';
  if (detEl) detEl.textContent = cardDet ? cardDet.textContent.trim() : '';

  // Render cue
  if (cueCard && cueText) {
    if (origCue) {
      var rawCue = origCue.textContent.trim();
      cueText.textContent = rawCue;
      cueCard.style.display = rawCue ? 'flex' : 'none';
    } else {
      cueCard.style.display = 'none';
    }
  }

  // Render Video buttons if present
  if (videoArea) {
    var origVideos = card.querySelectorAll('.video-btn');
    if (origVideos.length > 0) {
      var vHtml = '';
      origVideos.forEach(function(vb, vIdx) {
        var oc = vb.getAttribute('onclick') || '';
        var vt = vb.textContent.replace(/^[▶►]\\s*/, '').trim() || ('Demo ' + (vIdx + 1));
        vHtml += '<button class="focus-video-btn" onclick="' + escapeFocusHtml(oc) + '"><svg viewBox="0 0 24 24" style="width:14px;height:14px;fill:currentColor;"><path d="M8 5v14l11-7z"/></svg> ' + escapeFocusHtml(vt) + '</button>';
      });
      videoArea.innerHTML = vHtml;
      videoArea.style.display = 'flex';
    } else {
      videoArea.innerHTML = '';
      videoArea.style.display = 'none';
    }
  }

  // Render Segmented Progress Track
  renderFocusProgressTrack();

  // Render Interactive Sets
  renderFocusSets(card);

  // Configure Rest Timer
  setupFocusTimer(card);

  // Configure Prev/Next buttons
  var prevBtn = document.getElementById('focus-prev-btn');
  var nextBtn = document.getElementById('focus-next-btn');

  if (prevBtn) {
    if (focusIndex === 0) {
      prevBtn.style.opacity = '0.3';
      prevBtn.style.pointerEvents = 'none';
    } else {
      prevBtn.style.opacity = '1';
      prevBtn.style.pointerEvents = 'auto';
    }
  }

  if (nextBtn) {
    if (focusIndex === focusCards.length - 1) {
      nextBtn.innerHTML = '<span>Finish Workout</span>';
      nextBtn.classList.add('btn-finish-mode');
    } else {
      nextBtn.innerHTML = '<span>Next Exercise →</span>';
      nextBtn.classList.remove('btn-finish-mode');
    }
  }
}

function renderFocusProgressTrack() {
  var track = document.getElementById('focus-progress-track');
  if (!track) return;
  var html = '';
  focusCards.forEach(function(c, i) {
    var cbs = c.querySelectorAll('.set-checkbox input[type="checkbox"]');
    var isDone = cbs.length > 0 && Array.from(cbs).every(function(box) { return box.checked; });
    var classes = 'focus-progress-seg';
    if (i === focusIndex) classes += ' active';
    else if (isDone || i < focusIndex) classes += ' completed';
    html += '<div class="' + classes + '" onclick="goToFocusExercise(' + i + ')" title="Exercise ' + (i + 1) + (isDone ? ' (Completed)' : '') + '"></div>';
  });
  track.innerHTML = html;
}

function renderFocusSets(card) {
  var list = document.getElementById('focus-sets-list');
  var counter = document.getElementById('focus-sets-counter');
  if (!list) return;

  var setRows = card.querySelectorAll('.set-row');
  if (setRows.length === 0) {
    list.innerHTML = '<div style="color:var(--text-muted);font-size:0.85em;text-align:center;padding:12px;">No tracked sets for this exercise.</div>';
    if (counter) counter.textContent = '';
    return;
  }

  var completedCount = 0;
  var html = '';
  setRows.forEach(function(row, idx) {
    var cb = row.querySelector('.set-checkbox input[type="checkbox"]');
    var wInput = row.querySelector('.set-input.set-weight');
    var rInput = row.querySelector('.set-input.set-reps');
    var labelSpan = row.querySelector('.set-checkbox span');
    var labelText = labelSpan ? labelSpan.textContent.trim() : ('Set ' + (idx + 1));
    var setId = cb ? cb.getAttribute('data-set-id') : idx;
    var isChecked = cb && cb.checked;
    if (isChecked) completedCount++;

    var wVal = wInput ? wInput.value : '';
    var rVal = rInput ? rInput.value : '';

    html += '<div class="focus-set-row ' + (isChecked ? 'completed' : '') + '" id="focus-row-' + setId + '">';
    html += '  <label class="focus-set-checkbox">';
    html += '    <input type="checkbox" data-set-id="' + setId + '" ' + (isChecked ? 'checked' : '') + ' onchange="syncFocusCheckbox(this)">';
    html += '    <span>' + escapeFocusHtml(labelText) + '</span>';
    html += '  </label>';
    html += '  <div class="focus-set-inputs">';
    html += '    <div class="focus-input-wrap">';
    html += '      <input type="text" class="focus-input focus-weight" data-set-id="' + setId + '" data-field="weight" placeholder="—" inputmode="decimal" value="' + escapeFocusHtml(wVal) + '" oninput="syncFocusInput(this)">';
    html += '      <span class="focus-input-unit">kg</span>';
    html += '    </div>';
    html += '    <div class="focus-input-wrap">';
    html += '      <input type="text" class="focus-input focus-reps" data-set-id="' + setId + '" data-field="reps" placeholder="—" inputmode="numeric" value="' + escapeFocusHtml(rVal) + '" oninput="syncFocusInput(this)">';
    html += '      <span class="focus-input-unit">reps</span>';
    html += '    </div>';
    html += '  </div>';
    html += '</div>';
  });

  list.innerHTML = html;
  if (counter) {
    counter.textContent = completedCount + ' / ' + setRows.length + ' Done';
    counter.style.color = (completedCount === setRows.length && setRows.length > 0) ? 'var(--teal)' : 'var(--text-muted)';
  }
}

function syncFocusCheckbox(el) {
  var setId = el.getAttribute('data-set-id');
  var checked = el.checked;
  var orig = document.querySelector('.set-checkbox input[data-set-id="' + setId + '"]');
  if (orig) {
    orig.checked = checked;
    if (typeof updateProgress === 'function') updateProgress();
  }
  var row = document.getElementById('focus-row-' + setId);
  if (row) {
    if (checked) row.classList.add('completed');
    else row.classList.remove('completed');
  }

  var card = focusCards[focusIndex];
  if (card) {
    var setRows = card.querySelectorAll('.set-row');
    var done = 0;
    setRows.forEach(function(r) {
      var c = r.querySelector('.set-checkbox input[type="checkbox"]');
      if (c && c.checked) done++;
    });
    var counter = document.getElementById('focus-sets-counter');
    if (counter) {
      counter.textContent = done + ' / ' + setRows.length + ' Done';
      counter.style.color = (done === setRows.length && setRows.length > 0) ? 'var(--teal)' : 'var(--text-muted)';
    }
    renderFocusProgressTrack();
  }

  if (checked && currentFocusRestSeconds > 0) {
    showToast('Set completed! Tap Rest (' + currentFocusRestSeconds + 's)');
  }
}

function syncFocusInput(el) {
  var setId = el.getAttribute('data-set-id');
  var type = el.getAttribute('data-field');
  var val = el.value;
  var orig = document.querySelector('.set-input.set-' + type + '[data-set-id="' + setId + '"]');
  if (orig) {
    orig.value = val;
    if (typeof saveInputs === 'function') saveInputs();
  }
}

function setupFocusTimer(card) {
  var restBadge = card.querySelector('.badge-rest');
  var cardDet = card.querySelector('.details');
  var secMatch = null;
  if (restBadge) {
    var oc = restBadge.getAttribute('onclick') || '';
    var parts = oc.split('startRestFromBadge(');
    if (parts.length > 1) {
      var n = parseInt(parts[1], 10);
      if (!isNaN(n) && n > 0) secMatch = n;
    }
    if (!secMatch) {
      var num = parseInt(restBadge.textContent.replace(/[^0-9]/g, ''), 10);
      if (!isNaN(num) && num > 0) secMatch = num;
    }
  }
  if (!secMatch && cardDet) {
    var dt = cardDet.textContent;
    var rIdx = dt.toLowerCase().indexOf('s rest');
    if (rIdx !== -1) {
      var sub = dt.slice(0, rIdx).trim();
      var num2 = parseInt(sub.replace(/[^0-9]/g, ''), 10);
      if (!isNaN(num2) && num2 > 0) secMatch = num2;
    }
  }
  currentFocusRestSeconds = secMatch || 60;

  var restLabel = document.getElementById('focus-rest-label');
  if (restLabel) {
    restLabel.textContent = currentFocusRestSeconds + 's Rest';
  }

  var fDisp = document.getElementById('focus-time-display');
  var m = Math.floor(timeLeft / 60).toString().padStart(2, '0');
  var s = (timeLeft % 60).toString().padStart(2, '0');
  if (fDisp) {
    fDisp.innerText = m + ':' + s;
    if (isRunning) fDisp.classList.add('active');
    else fDisp.classList.remove('active');
  }

  var fStart = document.getElementById('focus-start-btn');
  if (fStart) {
    fStart.innerHTML = isRunning ? ICON_PAUSE : ICON_PLAY;
    if (isRunning) fStart.classList.add('running');
    else fStart.classList.remove('running');
  }
}

function triggerFocusRest() {
  var sec = currentFocusRestSeconds || 60;
  startRestFromBadge(sec);
  var fDisp = document.getElementById('focus-time-display');
  if (fDisp) fDisp.classList.add('active');
  var fStart = document.getElementById('focus-start-btn');
  if (fStart) {
    fStart.innerHTML = ICON_PAUSE;
    fStart.classList.add('running');
  }
}

function finishSession() {
  if (typeof fireConfetti === 'function') fireConfetti();
  if (typeof showToast === 'function') showToast('Workout Completed! Outstanding effort!');
  if (typeof closeFocusMode === 'function') closeFocusMode();
  var shareBtn = document.getElementById('btn-recap');
  if (shareBtn) shareBtn.classList.add('btn-accent');
  var shareMenuItem = document.getElementById('menu-item-recap');
  if (shareMenuItem) shareMenuItem.classList.add('menu-item-accent');
  if (window.safeStorage) safeStorage.setItem(STORAGE_KEY + '_session_finished', 'true');
  setTimeout(function() {
    openClientModal('share-prompt-modal');
  }, 450);
}

function nextFocusExercise() {
  if (focusIndex < focusCards.length - 1) {
    focusIndex++;
    renderFocusExercise();
    var sb = document.getElementById('focus-scroll-body');
    if (sb) sb.scrollTop = 0;
  } else {
    finishSession();
  }
}

function prevFocusExercise() {
  if (focusIndex > 0) {
    focusIndex--;
    renderFocusExercise();
    var sb = document.getElementById('focus-scroll-body');
    if (sb) sb.scrollTop = 0;
  }
}

// Initial session state check & reset hook
(function() {
  function checkFinishedState() {
    try {
      if (window.safeStorage && safeStorage.getItem(STORAGE_KEY + '_session_finished') === 'true') {
        var shareBtn = document.getElementById('btn-recap');
        if (shareBtn) shareBtn.classList.add('btn-accent');
        var shareMenuItem = document.getElementById('menu-item-recap');
        if (shareMenuItem) shareMenuItem.classList.add('menu-item-accent');
      }
    } catch(e) {}
  }
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', checkFinishedState);
  } else {
    checkFinishedState();
  }

  var origConfirm = document.getElementById('modal-confirm-btn');
  if (origConfirm) {
    origConfirm.addEventListener('click', function() {
      var shareBtn = document.getElementById('btn-recap');
      if (shareBtn) shareBtn.classList.remove('btn-accent');
      var shareMenuItem = document.getElementById('menu-item-recap');
      if (shareMenuItem) shareMenuItem.classList.remove('menu-item-accent');
      if (window.safeStorage) safeStorage.removeItem(STORAGE_KEY + '_session_finished');
    });
  }
})();
`;


  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no, viewport-fit=cover">
<meta name="theme-color" content="#070A10">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
<meta name="referrer" content="strict-origin-when-cross-origin">
<title>${escapeHtml(clientName)} - Training Program</title>
<style>
${CLIENT_CSS_TEMPLATE}

/* Global Reset & Dark Theme Scrollbars */
*, *::before, *::after {
  box-sizing: border-box;
}

html {
  background-color: #070A10 !important;
  background: #070A10 !important;
  min-height: 100%;
}

body {
  margin: 0;
  padding: 0 0 86px 0 !important;
  width: 100%;
  min-height: 100vh;
  overflow-x: hidden !important;
  background-color: #070A10 !important;
  background: #070A10 !important;
  color: #F8FAFC !important;
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif !important;
  scrollbar-width: thin !important;
  scrollbar-color: rgba(255, 255, 255, 0.15) transparent !important;
  border-radius: 0 !important;
  clip-path: none !important;
  -webkit-clip-path: none !important;
}

::-webkit-scrollbar {
  width: 4px !important;
  height: 4px !important;
}
::-webkit-scrollbar-track {
  background: transparent !important;
}
::-webkit-scrollbar-thumb {
  background: rgba(255, 255, 255, 0.18) !important;
  border-radius: 4px !important;
}
::-webkit-scrollbar-thumb:hover {
  background: #00E5FF !important;
}

/* Header Layout & Safe Area Under iPhone Notch */
.header {
  padding: max(44px, env(safe-area-inset-top, 44px)) 16px 12px 16px !important;
  background: #0B0F17 !important;
  border-bottom: 1px solid rgba(255, 255, 255, 0.08) !important;
}

.header-top {
  display: flex !important;
  justify-content: space-between !important;
  align-items: center !important;
  gap: 8px !important;
}

.header-titles {
  flex: 1 !important;
  min-width: 110px !important;
}

.header-actions {
  display: flex !important;
  align-items: center !important;
  gap: 8px !important;
  flex-shrink: 0 !important;
}

.header-sub {
  font-size: 0.65em !important;
  color: #00E5FF !important;
  font-weight: 800 !important;
  letter-spacing: 1px !important;
  text-transform: uppercase !important;
  display: block !important;
}

h1 {
  font-size: 1.25em !important;
  font-weight: 900 !important;
  margin: 2px 0 0 0 !important;
  color: #FFFFFF !important;
  line-height: 1.15 !important;
}

/* Header Session Menu & Dropdown */
.session-menu-wrap {
  position: relative !important;
  display: inline-block !important;
}

.btn-session-menu {
  width: 38px !important;
  height: 38px !important;
  border-radius: 10px !important;
  background: rgba(255, 255, 255, 0.05) !important;
  border: 1px solid rgba(255, 255, 255, 0.12) !important;
  color: #F8FAFC !important;
  display: flex !important;
  align-items: center !important;
  justify-content: center !important;
  cursor: pointer !important;
  position: relative !important;
  transition: all 0.18s ease !important;
  padding: 0 !important;
  -webkit-tap-highlight-color: transparent !important;
}
.btn-session-menu:hover, .btn-session-menu:active, .btn-session-menu.active {
  background: rgba(0, 229, 255, 0.14) !important;
  border-color: rgba(0, 229, 255, 0.4) !important;
  color: #00E5FF !important;
}

.btn-session-menu .mui-icon {
  width: 20px !important;
  height: 20px !important;
  fill: currentColor !important;
}

.session-menu-badge {
  position: absolute !important;
  top: 7px !important;
  right: 7px !important;
  width: 7px !important;
  height: 7px !important;
  border-radius: 50% !important;
  background: #00E5FF !important;
  box-shadow: 0 0 8px #00E5FF !important;
  display: none !important;
  pointer-events: none !important;
}
.session-menu-badge.active {
  display: block !important;
  animation: pulse-badge 1.8s infinite ease-in-out !important;
}

@keyframes pulse-badge {
  0% { transform: scale(0.9); opacity: 0.8; }
  50% { transform: scale(1.3); opacity: 1; box-shadow: 0 0 12px #00E5FF; }
  100% { transform: scale(0.9); opacity: 0.8; }
}

.session-dropdown-menu {
  position: absolute !important;
  top: calc(100% + 8px) !important;
  right: 0 !important;
  min-width: 210px !important;
  background: rgba(13, 19, 31, 0.97) !important;
  backdrop-filter: blur(16px) !important;
  -webkit-backdrop-filter: blur(16px) !important;
  border: 1px solid rgba(255, 255, 255, 0.12) !important;
  border-radius: 14px !important;
  padding: 6px !important;
  box-shadow: 0 18px 40px rgba(0, 0, 0, 0.65), 0 0 0 1px rgba(0, 229, 255, 0.12) !important;
  z-index: 1000 !important;
  display: none !important;
  flex-direction: column !important;
  gap: 2px !important;
  transform-origin: top right !important;
  animation: sessionMenuIn 0.16s cubic-bezier(0.16, 1, 0.3, 1) !important;
}

.session-dropdown-menu.show {
  display: flex !important;
}

@keyframes sessionMenuIn {
  from { opacity: 0; transform: scale(0.94) translateY(-6px); }
  to { opacity: 1; transform: scale(1) translateY(0); }
}

.session-dropdown-item {
  display: flex !important;
  align-items: center !important;
  gap: 10px !important;
  padding: 10px 12px !important;
  border-radius: 9px !important;
  color: #F1F5F9 !important;
  font-size: 0.84em !important;
  font-weight: 600 !important;
  cursor: pointer !important;
  transition: 0.15s all ease !important;
  user-select: none !important;
  -webkit-tap-highlight-color: transparent !important;
}

.session-dropdown-item:hover, .session-dropdown-item:active {
  background: rgba(0, 229, 255, 0.12) !important;
  color: #00E5FF !important;
}

.session-dropdown-item .mui-icon {
  width: 16px !important;
  height: 16px !important;
  fill: currentColor !important;
  flex-shrink: 0 !important;
  opacity: 0.85 !important;
}
.session-dropdown-item:hover .mui-icon {
  opacity: 1 !important;
}

.session-dropdown-item.danger {
  color: #F87171 !important;
}
.session-dropdown-item.danger:hover, .session-dropdown-item.danger:active {
  background: rgba(239, 68, 68, 0.14) !important;
  color: #EF4444 !important;
}

.session-dropdown-item.menu-item-accent {
  background: rgba(0, 229, 255, 0.15) !important;
  color: #00E5FF !important;
  font-weight: 700 !important;
}

.session-dropdown-divider {
  height: 1px !important;
  background: rgba(255, 255, 255, 0.08) !important;
  margin: 4px 6px !important;
}

.btn-client-action {
  background: rgba(255, 255, 255, 0.05) !important;
  border: 1px solid rgba(255, 255, 255, 0.12) !important;
  color: #F8FAFC !important;
  font-size: 0.76em !important;
  font-weight: 700 !important;
  padding: 6px 10px !important;
  border-radius: 8px !important;
  cursor: pointer !important;
  display: inline-flex !important;
  align-items: center !important;
  gap: 5px !important;
  transition: 0.15s all ease !important;
  -webkit-tap-highlight-color: transparent !important;
  line-height: 1 !important;
  white-space: nowrap !important;
  flex-shrink: 0 !important;
}
.btn-client-action.btn-active-timer {
  background: rgba(0, 229, 255, 0.14) !important;
  border-color: rgba(0, 229, 255, 0.4) !important;
  color: #00E5FF !important;
}
.btn-client-action.btn-inactive-timer {
  background: rgba(255, 255, 255, 0.03) !important;
  border-color: rgba(255, 255, 255, 0.08) !important;
  color: rgba(255, 255, 255, 0.45) !important;
}

.session-notes-card {
  margin-top: 16px !important;
  margin-bottom: 0px !important;
}


  .btn-timer-icon {
    width: 36px;
    height: 36px;
    border-radius: 50%;
    background: var(--neumorph-surface);
    color: var(--text-muted);
    border: 1px solid rgba(255, 255, 255, 0.06);
    cursor: pointer;
    box-shadow: 2px 2px 5px var(--neumorph-dark), -2px -2px 5px var(--neumorph-light);
    display: inline-flex;
    align-items: center;
    justify-content: center;
    transition: 0.15s all ease;
    padding: 0;
    flex-shrink: 0;
  }
  .btn-timer-icon.btn-play {
    color: var(--purple);
  }
  .btn-timer-icon.btn-play.running {
    color: var(--teal);
    border-color: rgba(0, 229, 255, 0.35);
    box-shadow: inset 2px 2px 4px var(--neumorph-dark), inset -2px -2px 4px var(--neumorph-light);
  }
  .btn-timer-icon:active {
    transform: scale(0.94);
  }
  .btn-timer-badge {
    height: 32px;
    padding: 0 8px;
    border-radius: 16px;
    background: var(--neumorph-surface);
    color: var(--text-muted);
    border: 1px solid rgba(255, 255, 255, 0.05);
    font-size: 0.74em;
    font-weight: 700;
    cursor: pointer;
    box-shadow: 2px 2px 5px var(--neumorph-dark), -2px -2px 5px var(--neumorph-light);
    display: inline-flex;
    align-items: center;
    justify-content: center;
    flex-shrink: 0;
    transition: 0.15s all ease;
  }
  .btn-timer-badge:active {
    transform: scale(0.94);
    box-shadow: inset 1px 1px 3px var(--neumorph-dark);
  }
  .timer-btn-svg {
    width: 17px;
    height: 17px;
    fill: currentColor;
    display: block;
  }

.timer-bar {
  position: fixed !important;
  bottom: max(16px, env(safe-area-inset-bottom, 16px)) !important;
  left: 50% !important;
  transform: translateX(-50%) !important;
  width: calc(100% - 28px) !important;
  max-width: 410px !important;
  box-sizing: border-box !important;
  border-radius: 24px !important;
  transition: transform 0.28s cubic-bezier(0.16, 1, 0.3, 1), opacity 0.22s ease !important;
}
.timer-bar.hidden,
.timer-bar.keyboard-hidden {
  transform: translateX(-50%) translateY(120px) !important;
  opacity: 0 !important;
  pointer-events: none !important;
}

.btn-client-action:hover, .btn-client-action:active {
  background: rgba(0, 229, 255, 0.15) !important;
  border-color: rgba(0, 229, 255, 0.4) !important;
  transform: scale(0.97) !important;
}

.btn-client-action.btn-accent {
  background: rgba(0, 229, 255, 0.12) !important;
  border-color: rgba(0, 229, 255, 0.35) !important;
  color: #00E5FF !important;
}

.btn-client-action .mui-icon {
  width: 14px !important;
  height: 14px !important;
  fill: currentColor !important;
  flex-shrink: 0 !important;
}

/* Finish Workout Button */
.btn-finish-workout {
  width: 100% !important;
  margin-top: 20px !important;
  padding: 14px !important;
  background: linear-gradient(135deg, #00E5FF, #00A3FF) !important;
  color: #05080E !important;
  border: none !important;
  border-radius: 12px !important;
  font-size: 1em !important;
  font-weight: 800 !important;
  cursor: pointer !important;
  box-shadow: 0 4px 20px rgba(0, 229, 255, 0.3) !important;
  display: flex !important;
  align-items: center !important;
  justify-content: center !important;
  gap: 8px !important;
  transition: 0.15s all !important;
}
.btn-finish-workout:active {
  transform: scale(0.98) !important;
}

/* Modal Overlay & Box (Dark Glassmorphic) */
svg.mui-icon {
  width: 18px !important;
  height: 18px !important;
  fill: currentColor !important;
  flex-shrink: 0 !important;
}

.modal-overlay {
  position: fixed !important;
  top: 0 !important; left: 0 !important; right: 0 !important; bottom: 0 !important;
  background: rgba(4, 7, 12, 0.85) !important;
  backdrop-filter: blur(8px) !important;
  -webkit-backdrop-filter: blur(8px) !important;
  z-index: 1500 !important;
  display: flex !important;
  align-items: center !important;
  justify-content: center !important;
  opacity: 0 !important;
  pointer-events: none !important;
  transition: opacity 0.2s ease, transform 0.2s ease !important;
  padding: 46px 12px 20px 12px !important;
  overflow-y: auto !important;
}
.modal-overlay.show, .modal-overlay.active {
  opacity: 1 !important;
  pointer-events: auto !important;
}
.modal-overlay .modal-box {
  background: #0E1422 !important;
  border: 1px solid rgba(255, 255, 255, 0.12) !important;
  border-radius: 18px !important;
  padding: 16px 14px !important;
  width: 100% !important;
  max-width: 340px !important;
  box-shadow: 0 20px 50px rgba(0, 0, 0, 0.8) !important;
  transform: scale(0.95) translateY(10px) !important;
  transition: transform 0.2s ease !important;
  color: #F8FAFC !important;
  margin: auto !important;
}
.modal-overlay.show .modal-box, .modal-overlay.active .modal-box {
  transform: scale(1) translateY(0) !important;
}


/* Custom Client Dropdown (Glassmorphic) */
.client-dropdown {
  position: relative;
  width: 155px;
  flex-shrink: 0;
}
.client-dropdown-trigger {
  width: 100%;
  background: #070A10;
  border: 1px solid rgba(255, 255, 255, 0.15);
  border-radius: 8px;
  padding: 8px 10px;
  color: #FFF;
  font-size: 0.78em;
  font-weight: 700;
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 6px;
  cursor: pointer;
  transition: 0.18s all ease;
  user-select: none;
}
.client-dropdown-trigger:hover,
.client-dropdown-trigger.active {
  border-color: #00E5FF !important;
  box-shadow: 0 0 10px rgba(0, 229, 255, 0.25);
  background: #0B111D;
}
.client-dropdown-trigger .chevron-icon {
  fill: #94A3B8;
  transition: transform 0.2s ease;
}
.client-dropdown-trigger.active .chevron-icon {
  transform: rotate(180deg);
  fill: #00E5FF;
}
.client-dropdown-menu {
  position: absolute;
  top: calc(100% + 5px);
  left: 0;
  right: 0;
  background: #0D121D;
  border: 1px solid rgba(0, 229, 255, 0.35);
  border-radius: 10px;
  box-shadow: 0 10px 30px rgba(0, 0, 0, 0.8), 0 0 12px rgba(0, 229, 255, 0.2);
  padding: 4px;
  z-index: 100;
  display: none;
  flex-direction: column;
  gap: 3px;
}
.client-dropdown-menu.show {
  display: flex;
}
.client-dropdown-item {
  padding: 7px 10px;
  border-radius: 6px;
  color: #CBD5E1;
  font-size: 0.78em;
  font-weight: 700;
  display: flex;
  align-items: center;
  justify-content: space-between;
  cursor: pointer;
  transition: 0.15s all ease;
}
.client-dropdown-item:hover {
  background: rgba(0, 229, 255, 0.12);
  color: #FFF;
}
.client-dropdown-item.active {
  background: rgba(0, 229, 255, 0.16);
  color: #00E5FF;
}
.client-dropdown-item .check-icon {
  fill: #00E5FF;
}

/* Barbell Plate Chips */
.plate-chip {
  display: inline-flex !important;
  align-items: center !important;
  justify-content: center !important;
  font-weight: 800 !important;
  font-size: 0.72em !important;
  border-radius: 6px !important;
  padding: 4px 7px !important;
  margin: 3px !important;
  color: #05080E !important;
}
.plate-20 { background: #3B82F6 !important; color: #FFF !important; }
.plate-15 { background: #EAB308 !important; color: #000 !important; }
.plate-10 { background: #10B981 !important; color: #FFF !important; }
.plate-5 { background: #F8FAFC !important; color: #000 !important; }
.plate-2_5 { background: #EF4444 !important; color: #FFF !important; }
.plate-1_25 { background: #94A3B8 !important; color: #000 !important; }

/* Focus Mode Fullscreen Overlay (Cockpit Upgrade) */
body.focus-open {
  overflow: hidden !important;
}

.focus-overlay {
  position: fixed !important;
  inset: 0 !important;
  width: 100vw !important;
  height: 100vh !important;
  background: #070B12 !important;
  z-index: 1200 !important;
  display: none !important;
  align-items: center !important;
  justify-content: center !important;
  box-sizing: border-box !important;
  padding: 0 !important;
  margin: 0 !important;
}

.focus-overlay.show, .focus-overlay.active {
  display: flex !important;
}

.focus-container {
  width: 100% !important;
  max-width: 420px !important;
  height: 100% !important;
  display: flex !important;
  flex-direction: column !important;
  padding: calc(env(safe-area-inset-top, 14px) + 10px) 14px calc(env(safe-area-inset-bottom, 14px) + 12px) 14px !important;
  box-sizing: border-box !important;
}

.focus-header {
  display: flex !important;
  justify-content: space-between !important;
  align-items: center !important;
  margin-bottom: 8px !important;
  flex-shrink: 0 !important;
}

.focus-progress-badge {
  font-size: 0.82em !important;
  font-weight: 800 !important;
  color: var(--teal, #00E5FF) !important;
  letter-spacing: 0.5px !important;
  display: flex !important;
  align-items: center !important;
  gap: 6px !important;
}

.focus-btn-exit {
  background: rgba(255, 255, 255, 0.06) !important;
  border: 1px solid rgba(255, 255, 255, 0.12) !important;
  color: #94A3B8 !important;
  border-radius: 20px !important;
  padding: 5px 12px !important;
  font-size: 0.78em !important;
  font-weight: 700 !important;
  cursor: pointer !important;
  display: inline-flex !important;
  align-items: center !important;
  gap: 5px !important;
  transition: all 0.15s ease !important;
}
.focus-btn-exit:hover, .focus-btn-exit:active {
  background: rgba(255, 255, 255, 0.12) !important;
  color: #FFF !important;
}

.focus-progress-track {
  display: flex !important;
  gap: 5px !important;
  width: 100% !important;
  margin-bottom: 12px !important;
  flex-shrink: 0 !important;
}

.focus-progress-seg {
  flex: 1 !important;
  height: 4px !important;
  border-radius: 2px !important;
  background: rgba(255, 255, 255, 0.08) !important;
  cursor: pointer !important;
  transition: all 0.2s ease !important;
}
.focus-progress-seg.completed {
  background: #A78BFA !important;
  opacity: 0.85 !important;
}
.focus-progress-seg.active {
  background: var(--teal, #00E5FF) !important;
  box-shadow: 0 0 8px rgba(0, 229, 255, 0.6) !important;
  transform: scaleY(1.4) !important;
}

.focus-scroll-body {
  flex: 1 !important;
  overflow-y: auto !important;
  display: flex !important;
  flex-direction: column !important;
  gap: 12px !important;
  padding-bottom: 12px !important;
  -webkit-overflow-scrolling: touch !important;
}
.focus-scroll-body::-webkit-scrollbar {
  width: 4px;
}
.focus-scroll-body::-webkit-scrollbar-thumb {
  background: rgba(255, 255, 255, 0.15);
  border-radius: 2px;
}

.focus-hero-card {
  background: linear-gradient(180deg, rgba(18, 24, 38, 0.9) 0%, rgba(13, 18, 29, 0.95) 100%) !important;
  border: 1px solid rgba(255, 255, 255, 0.08) !important;
  border-top: 1px solid rgba(255, 255, 255, 0.16) !important;
  border-radius: 18px !important;
  padding: 16px 14px !important;
  box-shadow: 0 10px 28px rgba(0,0,0,0.4) !important;
  display: flex !important;
  flex-direction: column !important;
  gap: 8px !important;
}

.focus-category-badge {
  font-size: 0.72em !important;
  font-weight: 800 !important;
  text-transform: uppercase !important;
  color: #A78BFA !important;
  background: rgba(167, 139, 250, 0.12) !important;
  border: 1px solid rgba(167, 139, 250, 0.25) !important;
  padding: 3px 8px !important;
  border-radius: 10px !important;
  align-self: flex-start !important;
  letter-spacing: 0.5px !important;
}

.focus-title {
  font-size: 1.35em !important;
  font-weight: 800 !important;
  color: #F8FAFC !important;
  margin: 0 !important;
  line-height: 1.25 !important;
}

.focus-details-row {
  font-size: 0.88em !important;
  font-weight: 600 !important;
  color: #94A3B8 !important;
}

.focus-video-area {
  display: flex !important;
  gap: 8px !important;
  flex-wrap: wrap !important;
  margin-top: 2px !important;
}

.focus-video-btn {
  display: inline-flex !important;
  align-items: center !important;
  gap: 6px !important;
  background: rgba(0, 229, 255, 0.1) !important;
  border: 1px solid rgba(0, 229, 255, 0.3) !important;
  color: var(--teal, #00E5FF) !important;
  padding: 6px 12px !important;
  border-radius: 10px !important;
  font-size: 0.8em !important;
  font-weight: 700 !important;
  cursor: pointer !important;
  transition: all 0.15s ease !important;
}
.focus-video-btn:hover, .focus-video-btn:active {
  background: rgba(0, 229, 255, 0.2) !important;
  transform: scale(0.97) !important;
}

.focus-cue-card {
  background: rgba(255, 255, 255, 0.03) !important;
  border-left: 3px solid #A78BFA !important;
  border-radius: 0 10px 10px 0 !important;
  padding: 9px 12px !important;
  font-size: 0.82em !important;
  color: #CBD5E1 !important;
  font-style: italic !important;
  display: flex !important;
  gap: 8px !important;
  align-items: flex-start !important;
  line-height: 1.4 !important;
}

.focus-sets-card {
  background: rgba(13, 18, 29, 0.85) !important;
  border: 1px solid rgba(255, 255, 255, 0.06) !important;
  border-radius: 18px !important;
  padding: 14px !important;
  display: flex !important;
  flex-direction: column !important;
  gap: 8px !important;
}

.focus-sets-header {
  display: flex !important;
  justify-content: space-between !important;
  align-items: center !important;
  font-size: 0.76em !important;
  font-weight: 800 !important;
  color: #94A3B8 !important;
  letter-spacing: 0.5px !important;
  margin-bottom: 2px !important;
}

.focus-sets-list {
  display: flex !important;
  flex-direction: column !important;
  gap: 6px !important;
}

.focus-set-row {
  display: flex !important;
  align-items: center !important;
  justify-content: space-between !important;
  gap: 8px !important;
  padding: 8px 10px !important;
  border-radius: 12px !important;
  background: rgba(255, 255, 255, 0.02) !important;
  border: 1px solid rgba(255, 255, 255, 0.04) !important;
  transition: all 0.15s ease !important;
}
.focus-set-row.completed {
  background: rgba(0, 229, 255, 0.05) !important;
  border-color: rgba(0, 229, 255, 0.22) !important;
}

.focus-set-checkbox {
  display: flex !important;
  align-items: center !important;
  gap: 10px !important;
  cursor: pointer !important;
  font-size: 0.88em !important;
  font-weight: 700 !important;
  color: #F1F5F9 !important;
  user-select: none !important;
}

.focus-set-checkbox input[type="checkbox"] {
  appearance: none !important;
  -webkit-appearance: none !important;
  width: 22px !important;
  height: 22px !important;
  border: 2px solid rgba(255, 255, 255, 0.2) !important;
  border-radius: 6px !important;
  background: rgba(0, 0, 0, 0.3) !important;
  cursor: pointer !important;
  display: grid !important;
  place-content: center !important;
  transition: all 0.15s ease !important;
  margin: 0 !important;
}
.focus-set-checkbox input[type="checkbox"]:checked {
  background: var(--teal, #00E5FF) !important;
  border-color: var(--teal, #00E5FF) !important;
}
.focus-set-checkbox input[type="checkbox"]:checked::before {
  content: "" !important;
  width: 10px !important;
  height: 6px !important;
  border-left: 2.5px solid #070A10 !important;
  border-bottom: 2.5px solid #070A10 !important;
  transform: rotate(-45deg) translate(1px, -1px) !important;
}

.focus-set-inputs {
  display: flex !important;
  gap: 6px !important;
}

.focus-input-wrap {
  position: relative !important;
  display: flex !important;
  align-items: center !important;
}

.focus-input {
  width: 58px !important;
  height: 32px !important;
  background: rgba(0, 0, 0, 0.35) !important;
  border: 1px solid rgba(255, 255, 255, 0.1) !important;
  border-radius: 8px !important;
  color: #FFF !important;
  text-align: left !important;
  padding-left: 8px !important;
  padding-right: 22px !important;
  font-size: 0.84em !important;
  font-weight: 600 !important;
  outline: none !important;
  transition: border-color 0.15s !important;
  box-sizing: border-box !important;
}
.focus-input:focus {
  border-color: var(--teal, #00E5FF) !important;
  box-shadow: 0 0 8px rgba(0, 229, 255, 0.25) !important;
}

.focus-input-unit {
  position: absolute !important;
  right: 6px !important;
  font-size: 0.68em !important;
  color: #64748B !important;
  pointer-events: none !important;
  font-weight: 600 !important;
}

.focus-timer-card {
  background: rgba(14, 20, 32, 0.95) !important;
  border: 1px solid rgba(255, 255, 255, 0.08) !important;
  border-radius: 16px !important;
  padding: 10px 12px !important;
  display: flex !important;
  justify-content: space-between !important;
  align-items: center !important;
  gap: 6px !important;
}

.focus-timer-badge {
  height: 32px !important;
  padding: 0 10px !important;
  border-radius: 16px !important;
  background: var(--neumorph-surface, #121824) !important;
  color: var(--teal, #00E5FF) !important;
  border: 1px solid rgba(0, 229, 255, 0.25) !important;
  font-size: 0.76em !important;
  font-weight: 700 !important;
  cursor: pointer !important;
  display: inline-flex !important;
  align-items: center !important;
  gap: 4px !important;
  transition: all 0.15s ease !important;
}
.focus-timer-badge:active {
  transform: scale(0.95) !important;
}

.focus-timer-display {
  font-size: 1.4em !important;
  font-weight: 800 !important;
  font-variant-numeric: tabular-nums !important;
  letter-spacing: -0.5px !important;
  color: #F8FAFC !important;
  padding: 0 4px !important;
}
.focus-timer-display.active {
  color: var(--teal, #00E5FF) !important;
  text-shadow: 0 0 12px rgba(0, 229, 255, 0.5) !important;
}

.focus-timer-actions {
  display: flex !important;
  gap: 6px !important;
  align-items: center !important;
}

.focus-dock {
  display: flex !important;
  gap: 10px !important;
  padding-top: 8px !important;
  flex-shrink: 0 !important;
}

.focus-btn-prev {
  flex: 1 !important;
  padding: 13px !important;
  background: rgba(255, 255, 255, 0.05) !important;
  border: 1px solid rgba(255, 255, 255, 0.1) !important;
  border-radius: 14px !important;
  color: #CBD5E1 !important;
  font-weight: 700 !important;
  font-size: 0.9em !important;
  cursor: pointer !important;
  display: flex !important;
  align-items: center !important;
  justify-content: center !important;
  gap: 6px !important;
  transition: all 0.15s ease !important;
}
.focus-btn-prev:hover, .focus-btn-prev:active {
  background: rgba(255, 255, 255, 0.1) !important;
}

.focus-btn-next {
  flex: 1.6 !important;
  padding: 13px !important;
  background: linear-gradient(135deg, #00E5FF 0%, #0284C7 100%) !important;
  border: none !important;
  border-radius: 14px !important;
  color: #040810 !important;
  font-weight: 800 !important;
  font-size: 0.92em !important;
  cursor: pointer !important;
  display: flex !important;
  align-items: center !important;
  justify-content: center !important;
  gap: 6px !important;
  box-shadow: 0 4px 18px rgba(0, 229, 255, 0.35) !important;
  transition: all 0.15s ease !important;
}
.focus-btn-next:active {
  transform: scale(0.97) !important;
}
.focus-btn-next.btn-finish-mode {
  background: linear-gradient(135deg, #10B981 0%, #059669 100%) !important;
  color: #FFF !important;
  box-shadow: 0 4px 18px rgba(16, 185, 129, 0.4) !important;
}

/* Weekly Plan / Schedule Cards in Athlete Player */
.plan-container {
  display: flex !important;
  flex-direction: column !important;
  gap: 14px !important;
  padding: 16px 12px !important;
  width: 100% !important;
  max-width: 650px !important;
  margin: 0 auto !important;
}
.day-card {
  background: #111827 !important;
  border: 1px solid rgba(255, 255, 255, 0.08) !important;
  border-radius: 14px !important;
  padding: 16px 18px !important;
  box-shadow: 0 4px 16px rgba(0, 0, 0, 0.3) !important;
}
.day-name {
  font-size: 0.76em !important;
  font-weight: 800 !important;
  text-transform: uppercase !important;
  letter-spacing: 0.06em !important;
  color: #A78BFA !important;
  background: rgba(167, 139, 250, 0.12) !important;
  display: inline-block !important;
  padding: 3px 10px !important;
  border-radius: 6px !important;
  margin-bottom: 8px !important;
}
.day-focus {
  font-size: 1.15em !important;
  font-weight: 800 !important;
  color: #F8FAFC !important;
  margin-bottom: 8px !important;
  line-height: 1.3 !important;
}
.day-details {
  font-size: 0.88em !important;
  color: #94A3B8 !important;
  line-height: 1.6 !important;
  white-space: pre-line !important;
  background: rgba(0, 0, 0, 0.25) !important;
  border-radius: 8px !important;
  padding: 10px 12px !important;
  border-left: 3px solid #A78BFA !important;
}


/* iOS Safari Auto-Zoom Prevention */
@media screen and (max-width: 768px) {
  .set-input, input, select, textarea {
    font-size: 16px !important;
  }
}

</style>
</head>
<body>
<noscript>
  <div style="background:#EF4444; color:#FFF; padding:14px 16px; text-align:center; font-weight:700; font-size:0.9em; line-height:1.4;">
    Notice: JavaScript is blocked by Apple QuickLook preview. Tap the Share icon and choose "Open in Safari" to interact with workouts.
  </div>
</noscript>
<div class="header">
  <div class="header-top">
    <div class="header-titles">
      <span class="header-sub">Custom Fitness Plan</span>
      <h1>${escapeHtml(clientName)}</h1>
    </div>
    <div class="header-actions">
      <div class="session-menu-wrap">
        <button class="btn-session-menu" id="btn-session-menu" onclick="toggleSessionMenu(event)" aria-label="Session Options" title="Session Options">
          <svg class="mui-icon" viewBox="0 0 24 24"><path d="M12 8c1.1 0 2-.9 2-2s-.9-2-2-2-2 .9-2 2 .9 2 2 2zm0 2c-1.1 0-2 .9-2 2s.9 2 2 2 2-.9 2-2-.9-2-2-2zm0 6c-1.1 0-2 .9-2 2s.9 2 2 2 2-.9 2-2-.9-2-2-2z"/></svg>
          <span class="session-menu-badge" id="session-timer-badge"></span>
        </button>
        <div class="session-dropdown-menu" id="session-dropdown-menu">
          <div class="session-dropdown-item" id="menu-item-timer" onclick="toggleTimerVisibility(); closeSessionMenu();">
            <svg class="mui-icon" viewBox="0 0 24 24"><path d="M15 1H9v2h6V1zm-4 13h2V8h-2v6zm8.03-6.61l1.42-1.42c-.43-.51-.9-.99-1.41-1.41l-1.42 1.42C16.07 4.74 14.12 4 12 4c-4.97 0-9 4.03-9 9s4.02 9 9 9 9-4.03 9-9c0-2.12-.74-4.07-1.97-5.61zM12 20c-3.87 0-7-3.13-7-7s3.13-7 7-7 7 3.13 7 7-3.13 7-7 7z"/></svg>
            <span id="menu-timer-text">Rest Timer (On)</span>
          </div>
          <div class="session-dropdown-item" id="menu-item-recap" onclick="openRecapModal(); closeSessionMenu();">
            <svg class="mui-icon" viewBox="0 0 24 24"><path d="M18 16.08c-.76 0-1.44.3-1.96.77L8.91 12.7c.05-.23.09-.46.09-.7s-.04-.47-.09-.7l7.05-4.11c.54.5 1.25.81 2.04.81 1.66 0 3-1.34 3-3s-1.34-3-3-3-3 1.34-3 3c0 .24.04.47.09.7L8.04 9.81C7.5 9.31 6.79 9 6 9c-1.66 0-3 1.34-3 3s1.34 3 3 3c.79 0 1.5-.31 2.04-.81l7.12 4.16c-.05.21-.08.43-.08.65 0 1.61 1.31 2.92 2.92 2.92s2.92-1.31 2.92-2.92c0-1.61-1.31-2.92-2.92-2.92z"/></svg>
            <span>Share Workout</span>
          </div>
          <div class="session-dropdown-item" onclick="openPlateModal(); closeSessionMenu();">
            <svg class="mui-icon" viewBox="0 0 24 24"><path d="M20.57 14.86L22 13.43 20.57 12 17 15.57 8.43 7 12 3.43 10.57 2 9.14 3.43 7.71 2 5.57 4.14 4.14 2.71 2.71 4.14l1.43 1.43L2 7.71l1.43 1.43L7 5.57 15.57 14.14 12 17.71l1.43 1.43 1.43-1.43 1.43 1.43 2.14-2.14 1.43 1.43 1.43-1.43-1.43-1.43 1.43-1.43z"/></svg>
            <span>Plate Calculator</span>
          </div>
          <div class="session-dropdown-item" onclick="openFocusMode(); closeSessionMenu();">
            <svg class="mui-icon" viewBox="0 0 24 24"><path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm-2 14.5v-9l6 4.5-6 4.5z"/></svg>
            <span>Focus Mode</span>
          </div>
          <div class="session-dropdown-item" onclick="openHistoryModal(); closeSessionMenu();">
            <svg class="mui-icon" viewBox="0 0 24 24"><path d="M13 3c-4.97 0-9 4.03-9 9H1l3.89 3.89.07.14L9 12H6c0-3.87 3.13-7 7-7s7 3.13 7 7-3.13 7-7 7c-1.93 0-3.68-.79-4.94-2.06l-1.42 1.42C8.27 19.99 10.51 21 13 21c4.97 0 9-4.03 9-9s-4.03-9-9-9zm-1 5v5l4.28 2.54.72-1.21-3.5-2.08V8H12z"/></svg>
            <span>Workout History</span>
          </div>
          <div class="session-dropdown-divider"></div>
          <div class="session-dropdown-item danger" onclick="promptResetSession(); closeSessionMenu();">
            <svg class="mui-icon" viewBox="0 0 24 24"><path d="M17.65 6.35A7.958 7.958 0 0 0 12 4c-4.42 0-7.99 3.58-7.99 8s3.57 8 7.99 8c3.73 0 6.84-2.55 7.73-6h-2.08A5.99 5.99 0 0 1 12 18c-3.31 0-6-2.69-6-6s2.69-6 6-6c1.66 0 3.14.69 4.22 1.78L13 11h7V4l-2.35 2.35z"/></svg>
            <span>Reset Session</span>
          </div>
        </div>
      </div>
    </div>
  </div>

  <div class="nav-menu" id="nav-menu">
${navHtml}
  </div>

  <div class="progress-bar-container" id="progress-container">
    <span id="progress-text">0 / ${totalProgramSets} Sets Done</span>
    <div class="progress-track">
      <div class="progress-fill" id="progress-fill"></div>
    </div>
  </div>
</div>

<div class="container" id="content-container">
${pagesHtml}
</div>

<div class="timer-bar" id="timer-bar">
  <div class="timer-controls">
    <button class="btn-timer-badge" onclick="addTime(-15)" title="Subtract 15s">-15s</button>
    <button class="btn-timer-badge" onclick="addTime(30)" title="Add 30s">+30s</button>
  </div>
  <div class="timer-display" id="time-display">00:00</div>
  <div class="timer-controls">
    <button class="btn-timer-icon btn-play" id="start-btn" onclick="toggleTimer()" title="Start timer">
      <svg class="timer-btn-svg" viewBox="0 0 24 24"><path d="M8 5v14l11-7z"/></svg>
    </button>
    <button class="btn-timer-icon" onclick="resetTimer()" title="Reset timer">
      <svg class="timer-btn-svg" viewBox="0 0 24 24" style="width:16px;height:16px;"><path d="M12 5V1L7 6l5 5V7c3.31 0 6 2.69 6 6s-2.69 6-6 6-6-2.69-6-6H4c0 4.42 3.58 8 8 8s8-3.58 8-8-3.58-8-8-8z"/></svg>
    </button>
  </div>
</div>

<div id="toast" class="toast-notification">Rest complete! Next set!</div>

<div id="custom-modal" class="modal-overlay" onclick="onModalBackdropClick(event)">
  <div class="modal-box">
    <div class="modal-title" id="modal-title">Reset Sets</div>
    <div class="modal-msg" id="modal-msg">Are you sure you want to reset your tracked sets?</div>
    <div class="modal-actions">
      <button class="modal-btn-cancel" onclick="closeModal()">Cancel</button>
      <button class="modal-btn-confirm" id="modal-confirm-btn">Reset</button>
    </div>
  </div>
</div>

<div id="video-modal" class="modal-overlay" onclick="onVideoBackdropClick(event)">
  <div class="modal-box video-modal-box">
    <div class="video-modal-header">
      <span class="video-modal-title" id="video-title">Exercise Demo</span>
      <button class="modal-close-btn" onclick="closeVideoModal()"><svg class="mui-icon" viewBox="0 0 24 24" style="width:18px;height:18px;"><path d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z"/></svg></button>
    </div>
    <div class="video-player-container">
      <iframe id="video-iframe" src="" allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture" allowfullscreen referrerpolicy="strict-origin-when-cross-origin"></iframe>
    </div>
        <div class="video-modal-footer">
      <a id="video-external-link" class="btn-open-youtube" href="#" target="_blank" rel="noopener noreferrer">
        <svg viewBox="0 0 24 24" style="width:18px;height:18px;fill:currentColor;"><path d="M10 15l5.19-3L10 9v6m11.56-7.83c.13.47.22 1.1.28 1.9.07.8.1 1.49.1 2.09L22 12c0 2.19-.16 3.8-.44 4.83-.25.9-.83 1.48-1.73 1.73-.47.13-1.33.22-2.65.28-1.3.07-2.49.1-3.59.1L12 22c-4.19 0-6.8-.16-7.83-.44-.9-.25-1.48-.83-1.73-1.73-.13-.47-.22-1.1-.28-1.9-.07-.8-.1-1.49-.1-2.09L2 12c0-2.19.16-3.8.44-4.83.25-.9.83-1.48 1.73-1.73.47-.13 1.33-.22 2.65-.28 1.3-.07 2.49-.1 3.59-.1L12 2c4.19 0 6.8.16 7.83.44.9.25 1.48.83 1.73 1.73z"/></svg>
        <span>Watch in YouTube App / Tab</span>
      </a>
      <div class="video-modal-hint">If playback is blocked in file mode, tap to watch directly.</div>
    </div>
  </div>
</div>


<!-- MODAL: SHARE WORKOUT PROMPT -->
<div id="share-prompt-modal" class="modal-overlay" onclick="if(event.target===this)closeClientModal('share-prompt-modal')">
  <div class="modal-box" style="text-align:center;">
    <div style="margin-bottom:8px;"><svg viewBox="0 0 24 24" style="width:36px;height:36px;fill:#00E5FF;"><path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm-2 15l-5-5 1.41-1.41L10 14.17l7.59-7.59L19 8l-9 9z"/></svg></div>
    <div class="modal-title" style="margin-bottom:6px; font-size:1.15em; font-weight:800; color:#F8FAFC;">Workout Completed!</div>
    <div class="modal-msg" style="margin-bottom:18px; line-height:1.45; color:#94A3B8; font-size:0.88em;">
      Outstanding effort! Would you like to share your workout recap with your coach?
    </div>
    <div class="modal-actions" style="display:flex; gap:10px;">
      <button class="modal-btn-cancel" onclick="closeClientModal('share-prompt-modal')" style="flex:1;">Done</button>
      <button class="btn-finish-workout" onclick="closeClientModal('share-prompt-modal'); openRecapModal();" style="margin:0; flex:1.4; padding:10px 12px; font-size:0.9em;">
        <svg class="mui-icon" viewBox="0 0 24 24" style="width:16px;height:16px;"><path d="M18 16.08c-.76 0-1.44.3-1.96.77L8.91 12.7c.05-.23.09-.46.09-.7s-.04-.47-.09-.7l7.05-4.11c.54.5 1.25.81 2.04.81 1.66 0 3-1.34 3-3s-1.34-3-3-3-3 1.34-3 3c0 .24.04.47.09.7L8.04 9.81C7.5 9.31 6.79 9 6 9c-1.66 0-3 1.34-3 3s1.34 3 3 3c.79 0 1.5-.31 2.04-.81l7.12 4.16c-.05.21-.08.43-.08.65 0 1.61 1.31 2.92 2.92 2.92s2.92-1.31 2.92-2.92c0-1.61-1.31-2.92-2.92-2.92z"/></svg> Share Workout
      </button>
    </div>
  </div>
</div>

<!-- MODAL: WHATSAPP / COACH RECAP -->
<div id="recap-modal" class="modal-overlay" onclick="if(event.target===this)closeClientModal('recap-modal')">
  <div class="modal-box">
    <div class="modal-title" style="display:flex; justify-content:space-between; align-items:center; margin-bottom:10px; font-size:1.05em; font-weight:800; color:#00E5FF;">
      <span style="display:inline-flex; align-items:center; gap:8px;">
        <svg class="mui-icon" viewBox="0 0 24 24" style="width:20px;height:20px;"><path d="M18 16.08c-.76 0-1.44.3-1.96.77L8.91 12.7c.05-.23.09-.46.09-.7s-.04-.47-.09-.7l7.05-4.11c.54.5 1.25.81 2.04.81 1.66 0 3-1.34 3-3s-1.34-3-3-3-3 1.34-3 3c0 .24.04.47.09.7L8.04 9.81C7.5 9.31 6.79 9 6 9c-1.66 0-3 1.34-3 3s1.34 3 3 3c.79 0 1.5-.31 2.04-.81l7.12 4.16c-.05.21-.08.43-.08.65 0 1.61 1.31 2.92 2.92 2.92s2.92-1.31 2.92-2.92c0-1.61-1.31-2.92-2.92-2.92z"/></svg>
        Share Workout
      </span>
      <button class="btn-client-action" onclick="closeClientModal('recap-modal')" style="padding:4px 8px;">
        <svg class="mui-icon" viewBox="0 0 24 24" style="width:16px;height:16px;"><path d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z"/></svg>
      </button>
    </div>
    <div style="font-size:0.8em; color:#94A3B8; margin-bottom:8px;">Summary ready to send to your coach:</div>
    <textarea id="recap-text" style="width:100%; min-height:140px; background:#070A10; border:1px solid rgba(255,255,255,0.1); border-radius:10px; padding:10px; color:#F8FAFC; font-family:monospace; font-size:0.8em; resize:vertical;"></textarea>
    <div style="display:flex; gap:8px; margin-top:12px;">
      <button class="btn-finish-workout" onclick="copyRecapToClipboard()" style="margin:0; flex:1; padding:10px 8px; font-size:0.85em;"><svg class="mui-icon" viewBox="0 0 24 24" style="width:15px;height:15px;"><path d="M16 1H4c-1.1 0-2 .9-2 2v14h2V3h12V1zm3 4H8c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h11c1.1 0 2-.9 2-2V7c0-1.1-.9-2-2-2zm0 16H8V7h11v14z"/></svg> Copy Summary</button>
      <button class="btn-finish-workout" onclick="openWhatsAppChat()" style="margin:0; flex:1; padding:10px 8px; font-size:0.85em; background:#25D366; color:#FFF;"><svg class="mui-icon" viewBox="0 0 24 24" style="width:15px;height:15px;"><path d="M20 2H4c-1.1 0-1.99.9-1.99 2L2 22l4-4h14c1.1 0 2-.9 2-2V4c0-1.1-.9-2-2-2z"/></svg> WhatsApp</button>
    </div>
  </div>
</div>

<!-- MODAL: BARBELL PLATE CALCULATOR -->
<div id="plate-modal" class="modal-overlay" onclick="if(event.target===this)closeClientModal('plate-modal')">
  <div class="modal-box">
    <div class="modal-title" style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px; font-size:1.05em; font-weight:800; color:#00E5FF;">
      <span style="display:inline-flex; align-items:center; gap:8px;">
        <svg class="mui-icon" viewBox="0 0 24 24" style="width:20px;height:20px;"><path d="M20.57 14.86L22 13.43 20.57 12 17 15.57 8.43 7 12 3.43 10.57 2 9.14 3.43 7.71 2 5.57 4.14 4.14 2.71 2.71 4.14l1.43 1.43L2 7.71l1.43 1.43L7 5.57 15.57 14.14 12 17.71l1.43 1.43 1.43-1.43 1.43 1.43 2.14-2.14 1.43 1.43 1.43-1.43-1.43-1.43 1.43-1.43z"/></svg>
        Plate Calculator
      </span>
      <button class="btn-client-action" onclick="closeClientModal('plate-modal')" style="padding:4px 8px;">
        <svg class="mui-icon" viewBox="0 0 24 24" style="width:16px;height:16px;"><path d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z"/></svg>
      </button>
    </div>
    
    <div style="margin-bottom:12px;">
      <label style="font-size:0.75em; color:#94A3B8; display:block; margin-bottom:4px; font-weight:700;">Target Weight (kg) & Barbell</label>
      <div style="display:flex; gap:8px; align-items:center;">
        <input type="number" id="plate-target-weight" value="80" step="2.5" style="flex:1; min-width:75px; background:#070A10; border:1px solid rgba(255,255,255,0.15); border-radius:8px; padding:8px 10px; color:#FFF; font-size:1.05em; font-weight:800;" oninput="calcPlates()">
        <div class="client-dropdown" id="plate-bar-dropdown">
          <button type="button" class="client-dropdown-trigger" id="plate-bar-trigger" onclick="togglePlateBarDropdown(event)">
            <span id="plate-bar-trigger-text">20kg (Olympic)</span>
            <svg class="mui-icon chevron-icon" viewBox="0 0 24 24" style="width:15px;height:15px;"><path d="M7 10l5 5 5-5z"/></svg>
          </button>
          <div class="client-dropdown-menu" id="plate-bar-menu">
            <div class="client-dropdown-item active" id="bar-opt-20" onclick="selectPlateBar(20, '20kg (Olympic)')">
              <span>20kg (Olympic)</span>
              <svg class="mui-icon check-icon" id="bar-chk-20" viewBox="0 0 24 24" style="width:14px;height:14px;"><path d="M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41z"/></svg>
            </div>
            <div class="client-dropdown-item" id="bar-opt-15" onclick="selectPlateBar(15, '15kg (Tech/Women)')">
              <span>15kg (Tech/Women)</span>
              <svg class="mui-icon check-icon" id="bar-chk-15" viewBox="0 0 24 24" style="width:14px;height:14px;display:none;"><path d="M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41z"/></svg>
            </div>
          </div>
          <input type="hidden" id="plate-bar-weight" value="20">
        </div>
      </div>
    </div>

    <div style="background:#070A10; border:1px solid rgba(255,255,255,0.08); border-radius:10px; padding:10px 12px; margin-bottom:10px;">
      <div style="font-size:0.72em; color:#94A3B8; text-transform:uppercase; font-weight:800;">Plates per Side:</div>
      <div id="plate-result-chips" style="min-height:30px; display:flex; align-items:center; flex-wrap:wrap; margin-top:4px;"></div>
      <div id="plate-side-weight" style="font-size:0.75em; color:#00E5FF; margin-top:4px; font-weight:700;"></div>
    </div>

    <div style="background:#070A10; border:1px solid rgba(255,255,255,0.08); border-radius:10px; padding:10px 12px;">
      <div style="font-size:0.72em; color:#94A3B8; text-transform:uppercase; font-weight:800; margin-bottom:6px;">Warm-Up Ramp Protocol:</div>
      <div id="plate-warmup-list" style="font-size:0.78em; color:#F8FAFC; display:flex; flex-direction:column; gap:3px;"></div>
    </div>
  </div>
</div>

<!-- MODAL: WORKOUT HISTORY -->
<div id="history-modal" class="modal-overlay" onclick="if(event.target===this)closeClientModal('history-modal')">
  <div class="modal-box">
    <div class="modal-title" style="display:flex; justify-content:space-between; align-items:center; margin-bottom:10px; font-size:1.05em; font-weight:800; color:#00E5FF;">
      <span style="display:inline-flex; align-items:center; gap:8px;">
        <svg class="mui-icon" viewBox="0 0 24 24" style="width:20px;height:20px;"><path d="M13 3c-4.97 0-9 4.03-9 9H1l3.89 3.89.07.14L9 12H6c0-3.87 3.13-7 7-7s7 3.13 7 7-3.13 7-7 7c-1.93 0-3.68-.79-4.94-2.06l-1.42 1.42C8.27 19.99 10.51 21 13 21c4.97 0 9-4.03 9-9s-4.03-9-9-9zm-1 5v5l4.28 2.54.72-1.21-3.5-2.08V8H12z"/></svg>
        Workout History
      </span>
      <button class="btn-client-action" onclick="closeClientModal('history-modal')" style="padding:4px 8px;">
        <svg class="mui-icon" viewBox="0 0 24 24" style="width:16px;height:16px;"><path d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z"/></svg>
      </button>
    </div>
    <div id="history-log-list" style="max-height:260px; overflow-y:auto; margin:10px 0; display:flex; flex-direction:column; gap:6px;">
      <div style="color:#94A3B8; font-size:0.8em; text-align:center; padding:16px;">No workout sessions logged yet. Complete a workout and tap "Finish & Log Workout" to save.</div>
    </div>
  </div>
</div>

<!-- OVERLAY: FULL-SCREEN FOCUS WORKOUT COCKPIT -->
<div id="focus-overlay" class="focus-overlay">
  <div class="focus-container">
    <div class="focus-header">
      <div class="focus-progress-badge" id="focus-progress-badge">
        <svg class="mui-icon" viewBox="0 0 24 24" style="width:16px;height:16px;"><path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm-1 14H9v-2h2v2zm0-4H9V7h2v5z"/></svg>
        <span id="focus-badge-text">EXERCISE 1 OF 4</span>
      </div>
      <button class="focus-btn-exit" onclick="closeFocusMode()" title="Exit Focus Mode">
        <svg viewBox="0 0 24 24" style="width:14px;height:14px;fill:currentColor;"><path d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z"/></svg> Exit Focus
      </button>
    </div>

    <div class="focus-progress-track" id="focus-progress-track"></div>

    <div class="focus-scroll-body" id="focus-scroll-body">
      <!-- Hero Details Card -->
      <div class="focus-hero-card">
        <div class="focus-category-badge" id="focus-ex-category">WARM-UP</div>
        <h2 class="focus-title" id="focus-ex-title">Exercise Title</h2>
        <div class="focus-details-row" id="focus-ex-details">Sets  -  Weight  -  Rest</div>
        <div class="focus-video-area" id="focus-video-area"></div>
        <div class="focus-cue-card" id="focus-cue-card">
          <svg viewBox="0 0 24 24" style="width:15px;height:15px;fill:#A78BFA;flex-shrink:0;"><path d="M9 21c0 .55.45 1 1 1h4c.55 0 1-.45 1-1v-1H9v1zm3-19C8.14 2 5 5.14 5 9c0 2.38 1.19 4.47 3 5.74V17c0 .55.45 1 1 1h6c.55 0 1-.45 1-1v-2.26c1.81-1.27 3-3.36 3-5.74 0-3.86-3.14-7-7-7z"/></svg>
          <span id="focus-cue-text">Coaching Cue</span>
        </div>
      </div>

      <!-- Interactive Sets Card -->
      <div class="focus-sets-card">
        <div class="focus-sets-header">
          <span>LOG SETS</span>
          <span id="focus-sets-counter">0 / 0 Done</span>
        </div>
        <div class="focus-sets-list" id="focus-sets-list"></div>
      </div>

      <!-- Embedded Rest Timer Card -->
      <div class="focus-timer-card" id="focus-timer-card">
        <button class="focus-timer-badge" id="focus-btn-quick-rest" onclick="triggerFocusRest()" title="Tap to start rest timer">
          <svg class="mui-icon" viewBox="0 0 24 24" style="width:14px;height:14px;"><path d="M15 1H9v2h6V1zm-4 13h2V8h-2v6zm8.03-6.61l1.42-1.42c-.43-.51-.9-.99-1.41-1.41l-1.42 1.42C16.07 4.74 14.12 4 12 4c-4.97 0-9 4.03-9 9s4.02 9 9 9 9-4.03 9-9c0-2.12-.74-4.07-1.97-5.61zM12 20c-3.87 0-7-3.13-7-7s3.13-7 7-7 7 3.13 7 7-3.13 7-7 7z"/></svg> <span id="focus-rest-label">60s Rest</span>
        </button>
        <div class="focus-timer-display" id="focus-time-display">00:00</div>
        <div class="focus-timer-actions">
          <button class="btn-timer-icon btn-play" id="focus-start-btn" onclick="toggleTimer()" title="Start timer">
            <svg class="timer-btn-svg" viewBox="0 0 24 24"><path d="M8 5v14l11-7z"/></svg>
          </button>
          <button class="btn-timer-icon" onclick="resetTimer()" title="Reset timer">
            <svg class="timer-btn-svg" viewBox="0 0 24 24" style="width:16px;height:16px;"><path d="M12 5V1L7 6l5 5V7c3.31 0 6 2.69 6 6s-2.69 6-6 6-6-2.69-6-6H4c0 4.42 3.58 8 8 8s8-3.58 8-8-3.58-8-8-8z"/></svg>
          </button>
        </div>
      </div>
    </div>

    <!-- Bottom Navigation Dock -->
    <div class="focus-dock">
      <button class="focus-btn-prev" id="focus-prev-btn" onclick="prevFocusExercise()">← Prev</button>
      <button class="focus-btn-next" id="focus-next-btn" onclick="nextFocusExercise()">Next Exercise →</button>
    </div>
  </div>
</div>

<canvas id="confetti-canvas"></canvas>

${'<script>'}
${jsCode}
${clientExtraScript}
${'<' + '/script>'}
</body>
</html>`;
}

// ==========================================================================
// CALENDAR & WORKOUT LOG CONTROLLER
// ==========================================================================

const MONTH_NAMES = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December'
];

function getAllCalendarLogs() {
  const all = [];
  studioAthletes.forEach(a => {
    if (a.logs && Array.isArray(a.logs)) {
      a.logs.forEach(l => {
        all.push({
          ...l,
          athleteId: a.id,
          athleteName: a.name,
          athleteInitials: a.initials || a.name.slice(0, 2).toUpperCase()
        });
      });
    }
  });
  return all;
}

function renderCalendarView() {
  const year = calendarCurrentDate.getFullYear();
  const month = calendarCurrentDate.getMonth();

  // Title
  const titleEl = document.getElementById('cal-month-title');
  if (titleEl) titleEl.textContent = `${MONTH_NAMES[month]} ${year}`;

  // Athlete Filter Select
  const filterSelect = document.getElementById('cal-athlete-filter');
  if (filterSelect) {
    const currentVal = calendarSelectedAthleteId;
    filterSelect.innerHTML = `
      <option value="all" ${currentVal === 'all' ? 'selected' : ''}>All Athletes</option>
      ${studioAthletes.map(a => `<option value="${a.id}" ${currentVal === a.id ? 'selected' : ''}>${escapeHtml(a.name)}</option>`).join('')}
    `;
    enhanceSelect(filterSelect);
  }

  // Get relevant logs
  let logs = getAllCalendarLogs();
  if (calendarSelectedAthleteId !== 'all') {
    logs = logs.filter(l => l.athleteId === calendarSelectedAthleteId);
  }

  // Update total logs badge
  const totalBadge = document.getElementById('calendar-total-logs-badge');
  if (totalBadge) totalBadge.textContent = `${logs.length} Log${logs.length === 1 ? '' : 's'}`;

  // Calculate stats for current month
  const monthPrefix = `${year}-${String(month + 1).padStart(2, '0')}`;
  const monthLogs = logs.filter(l => l.date && l.date.startsWith(monthPrefix));
  const completedCount = monthLogs.filter(l => (l.status || 'completed') === 'completed').length;
  const scheduledCount = monthLogs.filter(l => l.status === 'scheduled').length;

  const statsEl = document.getElementById('cal-month-stats');
  if (statsEl) {
    statsEl.textContent = `${completedCount} Completed  -  ${scheduledCount} Scheduled`;
  }

  // Calendar Grid Calculation (Monday = 0, Sunday = 6)
  const firstDayIndex = (new Date(year, month, 1).getDay() + 6) % 7;
  const daysInCurrentMonth = new Date(year, month + 1, 0).getDate();
  const daysInPrevMonth = new Date(year, month, 0).getDate();

  const grid = document.getElementById('calendar-days-grid');
  if (!grid) return;

  let gridHtml = '';
  const today = new Date();
  const todayStr = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}-${String(today.getDate()).padStart(2, '0')}`;

  // Padded Previous Month Days
  for (let i = firstDayIndex - 1; i >= 0; i--) {
    const d = daysInPrevMonth - i;
    const prevMonth = month === 0 ? 11 : month - 1;
    const prevYear = month === 0 ? year - 1 : year;
    const dateStr = `${prevYear}-${String(prevMonth + 1).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
    const dayLogs = logs.filter(l => l.date === dateStr);

    gridHtml += renderCalendarCellHtml(d, dateStr, dayLogs, true, dateStr === todayStr);
  }

  // Current Month Days
  for (let d = 1; d <= daysInCurrentMonth; d++) {
    const dateStr = `${year}-${String(month + 1).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
    const dayLogs = logs.filter(l => l.date === dateStr);

    gridHtml += renderCalendarCellHtml(d, dateStr, dayLogs, false, dateStr === todayStr);
  }

  // Padded Next Month Days to complete grid row (multiple of 7)
  const totalCells = firstDayIndex + daysInCurrentMonth;
  const remainingCells = (7 - (totalCells % 7)) % 7;
  for (let d = 1; d <= remainingCells; d++) {
    const nextMonth = month === 11 ? 0 : month + 1;
    const nextYear = month === 11 ? year + 1 : year;
    const dateStr = `${nextYear}-${String(nextMonth + 1).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
    const dayLogs = logs.filter(l => l.date === dateStr);

    gridHtml += renderCalendarCellHtml(d, dateStr, dayLogs, true, dateStr === todayStr);
  }

  grid.innerHTML = gridHtml;
}

function renderCalendarCellHtml(dayNum, dateStr, dayLogs, isOtherMonth, isToday) {
  const badgesHtml = dayLogs.map(l => {
    const statusClass = l.status || 'completed';
    return `
      <div class="cal-session-badge ${statusClass}" 
           onclick="event.stopPropagation(); openCalendarDayModal('${dateStr}', '${l.athleteId}', '${l.id}')"
           title="${escapeHtml(l.athleteName)}: ${escapeHtml(l.sessionTitle || 'Workout')} (${statusClass})">
        <span style="display:flex; align-items:center; gap:5px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;">
          <span class="cal-badge-ath-avatar">${escapeHtml(l.athleteInitials || l.athleteName[0])}</span>
          <span style="overflow:hidden; text-overflow:ellipsis;">${escapeHtml(l.sessionTitle || 'Workout')}</span>
        </span>
      </div>
    `;
  }).join('');

  return `
    <div class="calendar-day-cell ${isOtherMonth ? 'other-month' : ''} ${isToday ? 'today' : ''}" onclick="openCalendarDayModal('${dateStr}')">
      <div class="cal-day-header">
        <span class="cal-day-number">${dayNum}</span>
        <button type="button" class="cal-add-log-btn" onclick="event.stopPropagation(); openCalendarDayModal('${dateStr}')" title="Log workout on this day">
          + Log
        </button>
      </div>
      <div class="cal-badges-container">
        ${badgesHtml}
      </div>
    </div>
  `;
}

function navCalendarMonth(delta) {
  calendarCurrentDate.setMonth(calendarCurrentDate.getMonth() + delta);
  renderCalendarView();
}

function goToCalendarToday() {
  calendarCurrentDate = new Date();
  renderCalendarView();
}

function changeCalendarAthleteFilter(athleteId) {
  calendarSelectedAthleteId = athleteId;
  renderCalendarView();
}

// -------------------------------------------------------------
// CALENDAR DAY MODAL & LOGGING CONTROLLER
// -------------------------------------------------------------

function setCalLogStatus(status, btnEl) {
  document.getElementById('cal-log-status-val').value = status;
  const parent = btnEl.closest('.status-pill-group');
  if (parent) {
    parent.querySelectorAll('.status-pill-btn').forEach(b => b.classList.remove('active'));
    btnEl.classList.add('active');
  }
}

function toggleCalAssignBox(forceOpen = null) {
  const body = document.getElementById('cal-assign-box-body');
  const btn = document.getElementById('btn-toggle-cal-assign');
  if (!body || !btn) return;
  const shouldOpen = forceOpen !== null ? forceOpen : (body.style.display === 'none');
  body.style.display = shouldOpen ? 'flex' : 'none';
  btn.innerHTML = shouldOpen ? `${MUI.close} Collapse` : `${MUI.add} Expand`;
}

function resolveSessionExerciseDetails(athlete, sessionId, logExercises = []) {
  if (logExercises && Array.isArray(logExercises) && logExercises.length > 0) {
    return logExercises.map(e => ({
      exercise: e.exercise || 'Exercise',
      details: e.sets || '',
      weight: '',
      rest: '',
      cue: '',
      category: 'Logged Sets',
      isLoggedPerformance: true
    }));
  }

  if (athlete && athlete.pages && sessionId) {
    const page = athlete.pages.find(p => p.id === sessionId || p.id === sessionId) || (typeof sessionId === 'number' ? athlete.pages[sessionId] : null);
    if (page && page.cards && page.cards.length > 0) {
      return page.cards.map(c => ({
        exercise: c.exercise || 'Exercise',
        details: c.details || `${c.total_sets || 3} sets`,
        weight: c.weight || '',
        rest: c.rest || '',
        cue: c.cue || '',
        category: c.category || 'Primary Strength',
        isLoggedPerformance: false
      }));
    }
  }

  if (typeof sessionTemplates !== 'undefined' && Array.isArray(sessionTemplates)) {
    const t = sessionTemplates.find(tmpl => tmpl.id === sessionId || tmpl.title === sessionId);
    if (t && t.cards && t.cards.length > 0) {
      return t.cards.map(c => ({
        exercise: c.exercise || 'Exercise',
        details: c.details || `${c.total_sets || 3} sets`,
        weight: c.weight || '',
        rest: c.rest || '',
        cue: c.cue || '',
        category: c.category || 'Template',
        isLoggedPerformance: false
      }));
    }
  }

  return [];
}

function jumpToSessionBuilder(athleteId, sessionId) {
  const ath = studioAthletes.find(a => a.id === athleteId) || getActiveAthlete();
  if (!ath) return;

  closeModal('calendar-day-modal');

  let sessionIdx = 0;
  if (ath.pages && ath.pages.length > 0) {
    const foundIdx = ath.pages.findIndex(p => p.id === sessionId);
    if (foundIdx >= 0) sessionIdx = foundIdx;
  }

  navigateTo('builder', ath.id, sessionIdx);
  showToast(`Editing ${ath.name}'s workout routine in Builder`);
}

function updateCalendarLogStatus(logId, athleteId, newStatus, dateStr) {
  const ath = studioAthletes.find(a => a.id === athleteId);
  if (!ath || !ath.logs) return;

  const log = ath.logs.find(l => l.id === logId);
  if (!log) return;

  log.status = newStatus;
  saveAthletesToStorage();

  renderCalDayWorkoutsInspector(dateStr);
  if (currentView === 'calendar') {
    renderCalendarView();
  } else if (currentView === 'athlete') {
    renderAthleteDetail();
  }
  showToast(`Updated status to ${newStatus.toUpperCase()}`);
}

function openCalendarDayModal(dateStr = null, athleteId = null, highlightLogId = null) {
  const modal = document.getElementById('calendar-day-modal');
  if (!modal) return;

  const today = new Date();
  const todayStr = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}-${String(today.getDate()).padStart(2, '0')}`;
  const targetDate = dateStr || todayStr;
  const targetAthleteId = athleteId || (calendarSelectedAthleteId !== 'all' ? calendarSelectedAthleteId : (activeAthleteId || studioAthletes[0]?.id));

  document.getElementById('cal-log-date').value = targetDate;
  document.getElementById('cal-day-modal-title').textContent = `Workouts • ${targetDate}`;

  const athSelect = document.getElementById('cal-log-athlete');
  if (athSelect) {
    athSelect.innerHTML = studioAthletes.map(a => `
      <option value="${a.id}" ${a.id === targetAthleteId ? 'selected' : ''}>${escapeHtml(a.name)}</option>
    `).join('');
    enhanceSelect(athSelect);
  }

  onCalLogAthleteChange();

  const isFuture = targetDate > todayStr;
  const defaultStatus = isFuture ? 'scheduled' : 'completed';
  setCalLogStatus(defaultStatus, modal.querySelector(`.status-pill-btn.${defaultStatus}`));
  document.getElementById('cal-log-notes').value = '';

  renderCalDayWorkoutsInspector(targetDate, highlightLogId);

  openModal('calendar-day-modal');
}

function onCalLogAthleteChange() {
  const athId = document.getElementById('cal-log-athlete')?.value;
  const ath = studioAthletes.find(a => a.id === athId) || studioAthletes[0];
  const sessSelect = document.getElementById('cal-log-session');
  if (!sessSelect || !ath) return;

  let optionsHtml = '';

  if (ath.pages && ath.pages.length > 0) {
    optionsHtml += `<optgroup label="${escapeHtml(ath.name)}'s Workout Routines">`;
    ath.pages.forEach((p, idx) => {
      const cardCount = p.cards ? p.cards.length : 0;
      optionsHtml += `<option value="${p.id || 'page-' + idx}">#${idx + 1} ${escapeHtml(p.title || p.nav_title || 'Session')} (${cardCount} exercises)</option>`;
    });
    optionsHtml += `</optgroup>`;
  }

  if (typeof sessionTemplates !== 'undefined' && sessionTemplates.length > 0) {
    optionsHtml += `<optgroup label="Master Session Templates">`;
    sessionTemplates.forEach(t => {
      const cardCount = t.cards ? t.cards.length : 0;
      optionsHtml += `<option value="template:${t.id}">Template: ${escapeHtml(t.title)} (${cardCount} exercises)</option>`;
    });
    optionsHtml += `</optgroup>`;
  }

  optionsHtml += `<optgroup label="Other"><option value="custom">General / Custom Workout</option></optgroup>`;

  sessSelect.innerHTML = optionsHtml;
  enhanceSelect(sessSelect);
}

function renderCalDayWorkoutsInspector(dateStr, highlightLogId = null) {
  const container = document.getElementById('cal-day-workouts-container');
  const countBadge = document.getElementById('cal-day-modal-count-badge');
  if (!container) return;

  const dayLogs = getAllCalendarLogs().filter(l => l.date === dateStr);
  if (countBadge) {
    countBadge.textContent = `${dayLogs.length} Workout${dayLogs.length === 1 ? '' : 's'}`;
  }

  if (dayLogs.length === 0) {
    container.innerHTML = `
      <div style="text-align:center; padding:12px 10px; background:rgba(255,255,255,0.02); border:1px dashed var(--border); border-radius:10px; color:var(--text-dim); font-size:0.82em;">
        No workouts recorded on this date. Assign or log below:
      </div>
    `;
    toggleCalAssignBox(true);
    return;
  }

  toggleCalAssignBox(false);

  container.innerHTML = dayLogs.map(l => {
    const isHighlight = l.id === highlightLogId;
    const ath = studioAthletes.find(a => a.id === l.athleteId);
    const exercises = resolveSessionExerciseDetails(ath, l.sessionId, l.exercises);
    const statusClass = l.status || 'completed';

    const canOpenBuilder = ath && ath.pages && ath.pages.some(p => p.id === l.sessionId);

    let exercisesListHtml = '';
    if (exercises.length === 0) {
      exercisesListHtml = `
        <div style="font-size:0.82em; color:var(--text-dim); padding:8px 0; font-style:italic;">
          ${escapeHtml(l.exercisesSummary || 'Workout session - Open in Builder to view or add exercises.')}
        </div>
      `;
    } else {
      exercisesListHtml = exercises.map(ex => `
        <div class="cal-exercise-item">
          <div style="display:flex; flex-direction:column; gap:2px; flex:1; min-width:0;">
            <div style="display:flex; align-items:center; gap:8px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;">
              <span class="soc-badge" style="font-size:0.7em; padding:2px 6px;">${escapeHtml(ex.category)}</span>
              <strong style="font-size:0.88em; color:var(--text-main); overflow:hidden; text-overflow:ellipsis;">${escapeHtml(ex.exercise)}</strong>
            </div>
            ${ex.cue ? `<div class="cal-exercise-cue">${MUI.comment} <span>${escapeHtml(ex.cue)}</span></div>` : ''}
          </div>
          
          <div style="display:flex; align-items:center; gap:8px; font-size:0.82em; flex-shrink:0;">
            ${ex.details ? `<span style="color:var(--text-main); font-weight:700;">${escapeHtml(ex.details)}</span>` : ''}
            ${ex.weight ? `<span style="color:var(--teal); font-weight:700; background:rgba(0,229,255,0.08); border:1px solid rgba(0,229,255,0.2); padding:2px 6px; border-radius:4px;">${escapeHtml(ex.weight)}</span>` : ''}
            ${ex.rest ? `<span style="color:var(--text-dim); font-size:0.9em;">${escapeHtml(ex.rest)}</span>` : ''}
          </div>
        </div>
      `).join('');
    }

    return `
      <div class="cal-workout-card" style="${isHighlight ? 'border-color:var(--teal); box-shadow:0 0 12px var(--teal-glow);' : ''}">
        <!-- Top Header -->
        <div class="cal-workout-card-header">
          <div style="display:flex; align-items:center; gap:10px;">
            <div class="cal-badge-ath-avatar" style="width:30px; height:30px; font-size:12px; background:rgba(0,229,255,0.2); color:var(--teal); border:1px solid rgba(0,229,255,0.35);">
              ${escapeHtml(l.athleteInitials || l.athleteName[0])}
            </div>
            <div>
              <div style="font-size:1.02em; font-weight:800; color:var(--text-main);">${escapeHtml(l.athleteName)}</div>
              <div style="font-size:0.86em; font-weight:700; color:var(--teal);">${escapeHtml(l.sessionTitle || 'Workout Session')}</div>
            </div>
          </div>

          <div style="display:flex; align-items:center; gap:8px;">
            ${canOpenBuilder ? `
              <button type="button" class="btn-header" style="color:var(--teal); border-color:rgba(0,229,255,0.35); font-size:0.78em; padding:5px 12px;" onclick="jumpToSessionBuilder('${l.athleteId}', '${l.sessionId}')" title="Edit this workout in the Exercise Builder">
                ${MUI.edit} Edit in Builder
              </button>
            ` : ''}
            <button type="button" class="btn-header" style="color:var(--rose); padding:5px 8px; font-size:0.78em;" onclick="deleteCalendarLog('${l.id}', '${l.athleteId}', '${dateStr}')" title="Remove this workout log from Calendar">
              ${MUI.delete}
            </button>
          </div>
        </div>

        <!-- Status Toggle Row -->
        <div style="display:flex; align-items:center; justify-content:space-between; margin-bottom:12px; background:rgba(0,0,0,0.25); border:1px solid var(--border); border-radius:8px; padding:6px 10px;">
          <span style="font-size:0.78em; font-weight:700; color:var(--text-muted);">Status:</span>
          <div class="status-pill-group" style="width:270px;">
            <button type="button" class="status-pill-btn ${statusClass === 'completed' ? 'active completed' : ''}" onclick="updateCalendarLogStatus('${l.id}', '${l.athleteId}', 'completed', '${dateStr}')">Completed</button>
            <button type="button" class="status-pill-btn ${statusClass === 'scheduled' ? 'active scheduled' : ''}" onclick="updateCalendarLogStatus('${l.id}', '${l.athleteId}', 'scheduled', '${dateStr}')">Scheduled</button>
            <button type="button" class="status-pill-btn ${statusClass === 'missed' ? 'active missed' : ''}" onclick="updateCalendarLogStatus('${l.id}', '${l.athleteId}', 'missed', '${dateStr}')">Missed</button>
          </div>
        </div>

        <!-- Exercises Breakdown -->
        <div style="margin-bottom:8px;">
          <div style="font-size:0.78em; font-weight:800; color:var(--text-muted); text-transform:uppercase; letter-spacing:0.04em; margin-bottom:6px;">
            Assigned Routine & Exercises (${exercises.length}):
          </div>
          <div class="cal-exercises-list">
            ${exercisesListHtml}
          </div>
        </div>

        <!-- Coach Notes -->
        ${l.notes ? `
          <div style="background:rgba(255,255,255,0.02); border:1px solid var(--border); border-radius:8px; padding:8px 12px; font-size:0.8em; color:var(--text-dim); margin-top:8px;">
            <strong style="color:var(--text-muted);">Notes:</strong> "${escapeHtml(l.notes)}"
          </div>
        ` : ''}
      </div>
    `;
  }).join('');
}

function saveCalendarDayLog() {
  const dateVal = document.getElementById('cal-log-date')?.value;
  const athId = document.getElementById('cal-log-athlete')?.value;
  const sessSelectVal = document.getElementById('cal-log-session')?.value;
  const statusVal = document.getElementById('cal-log-status-val')?.value || 'completed';
  const notesVal = document.getElementById('cal-log-notes')?.value || '';

  if (!dateVal) {
    showToast('Please select a valid workout date');
    return;
  }

  const ath = studioAthletes.find(a => a.id === athId);
  if (!ath) {
    showToast('Athlete not found');
    return;
  }

  if (!ath.logs) ath.logs = [];

  let sessionTitle = 'Workout Session';
  let targetSessionId = sessSelectVal;
  let exercisesSummary = '';
  let exercises = [];

  if (sessSelectVal && sessSelectVal.startsWith('template:')) {
    const tmplId = sessSelectVal.replace('template:', '');
    const tmpl = (typeof sessionTemplates !== 'undefined') ? sessionTemplates.find(t => t.id === tmplId) : null;
    if (tmpl) {
      sessionTitle = tmpl.title;
      const newPageId = 'page-' + Date.now() + '-' + Math.random().toString(36).substr(2, 4);
      if (!ath.pages) ath.pages = [];
      ath.pages.push({
        id: newPageId,
        title: tmpl.title,
        nav_title: tmpl.title.slice(0, 15),
        type: 'exercises',
        cards: JSON.parse(JSON.stringify(tmpl.cards || []))
      });
      targetSessionId = newPageId;
      const count = tmpl.cards ? tmpl.cards.length : 0;
      exercisesSummary = `${count} exercise${count === 1 ? '' : 's'}`;
    }
  } else if (sessSelectVal && sessSelectVal !== 'custom' && ath.pages) {
    const page = ath.pages.find(p => p.id === sessSelectVal) || ath.pages[0];
    if (page) {
      sessionTitle = page.title || page.nav_title || 'Session';
      const count = page.cards ? page.cards.length : 0;
      exercisesSummary = `${count} exercise${count === 1 ? '' : 's'}`;
    }
  } else {
    sessionTitle = 'Custom Workout';
    exercisesSummary = notesVal.slice(0, 30) || 'Custom Session';
  }

  const newLog = {
    id: 'log-' + Date.now() + '-' + Math.random().toString(36).substr(2, 4),
    athleteId: ath.id,
    athleteName: ath.name,
    athleteInitials: ath.initials || ath.name.slice(0, 2).toUpperCase(),
    sessionId: targetSessionId,
    sessionTitle: sessionTitle,
    date: dateVal,
    status: statusVal,
    exercisesSummary: exercisesSummary,
    exercises: exercises,
    notes: notesVal.trim(),
    loggedAt: new Date().toISOString()
  };

  ath.logs.unshift(newLog);
  saveAthletesToStorage();

  renderCalDayWorkoutsInspector(dateVal, newLog.id);
  showToast(`Assigned ${sessionTitle} to ${ath.name} on ${dateVal}!`);

  if (currentView === 'calendar') {
    renderCalendarView();
  } else if (currentView === 'athlete') {
    renderAthleteDetail();
  }
}

function deleteCalendarLog(logId, athleteId, dateStr = null) {
  const ath = studioAthletes.find(a => a.id === athleteId);
  if (!ath || !ath.logs) return;

  ath.logs = ath.logs.filter(l => l.id !== logId);
  saveAthletesToStorage();
  showToast('Workout log removed');

  if (dateStr) {
    renderCalDayWorkoutsInspector(dateStr);
  }

  if (currentView === 'calendar') {
    renderCalendarView();
  } else if (currentView === 'athlete') {
    renderAthleteDetail();
  }
}

// -------------------------------------------------------------
// SMART WORKOUT RECAP IMPORTER
// -------------------------------------------------------------

function openImportRecapModal() {
  const textarea = document.getElementById('import-recap-text');
  if (textarea) textarea.value = '';

  const previewWrap = document.getElementById('import-recap-preview-wrap');
  if (previewWrap) {
    previewWrap.style.display = 'none';
    previewWrap.innerHTML = '';
  }

  const confirmBtn = document.getElementById('btn-confirm-import-recap');
  if (confirmBtn) {
    confirmBtn.disabled = true;
    confirmBtn.style.opacity = '0.45';
    confirmBtn.style.cursor = 'not-allowed';
  }

  parsedRecapCache = null;
  openModal('import-recap-modal');
}

function normalizeDateStringToISO(dateStr) {
  if (!dateStr) {
    const d = new Date();
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
  }
  const isoMatch = dateStr.match(new RegExp('^(\\d{4})-(\\d{1,2})-(\\d{1,2})$'));
  if (isoMatch) {
    return `${isoMatch[1]}-${String(isoMatch[2]).padStart(2, '0')}-${String(isoMatch[3]).padStart(2, '0')}`;
  }
  const slashMatch = dateStr.match(new RegExp('^(\\d{1,2})\\/(\\d{1,2})\\/(\\d{4})$'));
  if (slashMatch) {
    let partA = parseInt(slashMatch[1], 10);
    let partB = parseInt(slashMatch[2], 10);
    const y = slashMatch[3];
    if (partA > 12 && partB <= 12) {
      return `${y}-${String(partB).padStart(2, '0')}-${String(partA).padStart(2, '0')}`;
    }
  }
  const parsed = new Date(dateStr);
  if (!isNaN(parsed.getTime())) {
    return `${parsed.getFullYear()}-${String(parsed.getMonth() + 1).padStart(2, '0')}-${String(parsed.getDate()).padStart(2, '0')}`;
  }
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

function parseRecapText(text) {
  if (!text || !text.trim()) return null;

  const result = {
    athleteName: '',
    sessionTitle: '',
    date: '',
    exercises: [],
    notes: ''
  };

  const lines = text.split(String.fromCharCode(10)).map(l => l.trim()).filter(Boolean);

  // 1. Check for athlete name
  for (const line of lines) {
    const lower = line.toLowerCase();
    if (lower.startsWith('workout recap:')) {
      result.athleteName = line.slice(14).trim();
      break;
    } else if (lower.startsWith('athlete:')) {
      result.athleteName = line.slice(8).trim();
      break;
    }
  }

  // 2. Check for session title & date (e.g. "Session 1 - Upper Power (2026-09-06)")
  for (const line of lines) {
    const lower = line.toLowerCase();
    if (lower.startsWith('workout recap:') || lower.startsWith('athlete:') || lower.startsWith('coach notes:') || lower.startsWith('notes:') || lower.startsWith('---')) {
      continue;
    }
    const pOpen = line.lastIndexOf('(');
    const pClose = line.lastIndexOf(')');
    if (pOpen > 0 && pClose > pOpen && !line.includes(':')) {
      result.sessionTitle = line.slice(0, pOpen).trim();
      result.date = normalizeDateStringToISO(line.slice(pOpen + 1, pClose).trim());
      break;
    } else if ((lower.startsWith('session') || lower.startsWith('workout') || lower.startsWith('day')) && !line.includes(':')) {
      result.sessionTitle = line.trim();
      break;
    }
  }

  // Fallback date match if not found in session line
  if (!result.date) {
    for (const line of lines) {
      const parts = line.match(new RegExp('\\b\\d{4}-\\d{2}-\\d{2}\\b')) || line.match(new RegExp('\\b\\d{1,2}\\/\\d{1,2}\\/\\d{4}\\b'));
      if (parts) {
        result.date = normalizeDateStringToISO(parts[0]);
        break;
      }
    }
    if (!result.date) {
      result.date = normalizeDateStringToISO(null);
    }
  }

  // 3. Coach Notes
  for (const line of lines) {
    const lower = line.toLowerCase();
    if (lower.startsWith('coach notes:')) {
      result.notes = line.slice(12).trim();
      break;
    } else if (lower.startsWith('athlete notes:')) {
      result.notes = line.slice(14).trim();
      break;
    } else if (lower.startsWith('notes:')) {
      result.notes = line.slice(6).trim();
      break;
    }
  }

  // 4. Exercises & Sets (e.g. "Bent-Elbow Band Pull-Aparts: 2 sets (Light Band x 10, Light Band x 10)")
  for (const line of lines) {
    const lower = line.toLowerCase();
    if (line.startsWith('---') || lower.startsWith('workout recap:') || lower.startsWith('athlete:') || lower.startsWith('coach notes:') || lower.startsWith('notes:') || lower.startsWith('athlete notes:')) {
      continue;
    }
    const colonIdx = line.indexOf(':');
    if (colonIdx > 0) {
      const exName = line.slice(0, colonIdx).trim();
      const exSets = line.slice(colonIdx + 1).trim();
      // Validate that it has set/rep/weight keywords
      const sLow = exSets.toLowerCase();
      if (exName && exSets && (sLow.includes('set') || sLow.includes('x') || sLow.includes('kg') || sLow.includes('lb') || sLow.includes('band') || sLow.includes('bw') || sLow.includes('rep') || sLow.includes('done'))) {
        result.exercises.push({
          exercise: exName,
          sets: exSets
        });
      }
    }
  }

  // STRICT VALIDATION: Must have at least 1 parsed exercise OR (sessionTitle and recognized recap/session header)
  const hasExercises = result.exercises.length > 0;
  const hasValidSession = result.sessionTitle && (result.athleteName || text.toLowerCase().includes('workout recap') || text.toLowerCase().includes('session'));

  if (!hasExercises && !hasValidSession) {
    return null; // Reject gibberish like "cheese"
  }

  return result;
}

function previewImportedRecap() {
  const text = (document.getElementById('import-recap-text')?.value || '').trim();
  const previewWrap = document.getElementById('import-recap-preview-wrap');
  const confirmBtn = document.getElementById('btn-confirm-import-recap');
  if (!previewWrap) return;

  if (!text) {
    previewWrap.style.display = 'none';
    previewWrap.innerHTML = '';
    parsedRecapCache = null;
    if (confirmBtn) {
      confirmBtn.disabled = true;
      confirmBtn.style.opacity = '0.45';
      confirmBtn.style.cursor = 'not-allowed';
    }
    return;
  }

  const parsed = parseRecapText(text);

  if (!parsed) {
    // Show error preview warning
    parsedRecapCache = null;
    previewWrap.style.display = 'block';
    previewWrap.style.background = 'rgba(244, 63, 94, 0.08)';
    previewWrap.style.border = '1px solid rgba(244, 63, 94, 0.35)';
    previewWrap.innerHTML = `
      <div style="display:flex; align-items:center; gap:8px; color:var(--rose); font-size:0.85em; font-weight:700;">
        ${MUI.warning} No workout details recognized
      </div>
      <div style="font-size:0.8em; color:var(--text-muted); margin-top:4px; line-height:1.4;">
        The pasted text does not match a valid workout recap format. Please paste a recap containing a workout session name, date, and completed exercises.
      </div>
    `;

    if (confirmBtn) {
      confirmBtn.disabled = true;
      confirmBtn.style.opacity = '0.45';
      confirmBtn.style.cursor = 'not-allowed';
    }
    return;
  }

  // Valid parsed recap
  parsedRecapCache = parsed;
  previewWrap.style.display = 'block';
  previewWrap.style.background = 'rgba(0, 0, 0, 0.35)';
  previewWrap.style.border = '1px solid var(--border)';

  let matchedAth = null;
  if (parsed.athleteName) {
    const lower = parsed.athleteName.toLowerCase();
    matchedAth = studioAthletes.find(a => a.name.toLowerCase().includes(lower) || lower.includes(a.name.toLowerCase()));
  }
  if (!matchedAth) matchedAth = getActiveAthlete() || studioAthletes[0];

  const athOptionsHtml = studioAthletes.map(a => `
    <option value="${a.id}" ${matchedAth && a.id === matchedAth.id ? 'selected' : ''}>${escapeHtml(a.name)}</option>
  `).join('');

  const exercisesHtml = parsed.exercises.length > 0 ? parsed.exercises.map(e => `
    <div style="display:flex; justify-content:space-between; background:var(--bg-input); border:1px solid var(--border); border-radius:6px; padding:5px 8px; font-size:0.8em;">
      <strong style="color:var(--text-main);">${escapeHtml(e.exercise)}</strong>
      <span style="color:var(--teal); font-weight:700;">${escapeHtml(e.sets)}</span>
    </div>
  `).join('') : `<div style="font-size:0.8em; color:var(--text-dim);">General workout routine</div>`;

  previewWrap.innerHTML = `
    <div style="font-size:0.8em; font-weight:800; color:var(--teal); text-transform:uppercase; letter-spacing:0.04em; margin-bottom:10px; display:flex; align-items:center; gap:6px;">
      ${MUI.check} Parsed Workout Summary
    </div>

    <div style="display:grid; grid-template-columns:1fr 1fr; gap:10px; margin-bottom:10px;">
      <div>
        <label class="form-label" style="font-size:0.75em;">Matched Athlete</label>
        <select id="import-recap-ath-select" class="form-input" style="font-size:0.85em; padding:6px 10px;">
          ${athOptionsHtml}
        </select>
      </div>
      <div>
        <label class="form-label" style="font-size:0.75em;">Workout Date</label>
        <input type="date" id="import-recap-date-input" class="form-input" style="font-size:0.85em; padding:6px 10px;" value="${parsed.date}">
      </div>
    </div>

    <div style="margin-bottom:8px;">
      <div style="font-size:0.75em; color:var(--text-muted); font-weight:600;">Session:</div>
      <div id="import-recap-session-title" style="font-size:0.9em; font-weight:800; color:var(--text-main); margin-top:2px;">
        ${escapeHtml(parsed.sessionTitle || 'Workout Session')}
      </div>
    </div>

    <div style="margin-bottom:8px;">
      <div style="font-size:0.75em; color:var(--text-muted); font-weight:600;">Exercises Detected (${parsed.exercises.length}):</div>
      <div id="import-recap-exercises-list" style="margin-top:4px; max-height:100px; overflow-y:auto; display:flex; flex-direction:column; gap:4px;">
        ${exercisesHtml}
      </div>
    </div>

    ${parsed.notes ? `
      <div style="margin-top:6px;">
        <div style="font-size:0.75em; color:var(--text-muted); font-weight:600;">Athlete Notes:</div>
        <div style="font-size:0.82em; color:var(--text-dim); margin-top:2px; font-style:italic;">
          "${escapeHtml(parsed.notes)}"
        </div>
      </div>
    ` : ''}
  `;

  const newAthSelect = document.getElementById('import-recap-ath-select');
  if (newAthSelect) enhanceSelect(newAthSelect);

  if (confirmBtn) {
    confirmBtn.disabled = false;
    confirmBtn.style.opacity = '1';
    confirmBtn.style.cursor = 'pointer';
  }
}

function confirmImportRecap() {
  if (!parsedRecapCache) {
    const text = (document.getElementById('import-recap-text')?.value || '').trim();
    parsedRecapCache = parseRecapText(text);
  }

  if (!parsedRecapCache || (!parsedRecapCache.sessionTitle && parsedRecapCache.exercises.length === 0)) {
    showToast('Cannot import: No valid workout exercises or session found');
    return;
  }

  const athId = document.getElementById('import-recap-ath-select')?.value;
  const ath = studioAthletes.find(a => a.id === athId) || studioAthletes[0];
  if (!ath) {
    showToast('Selected athlete not found');
    return;
  }

  const dateVal = document.getElementById('import-recap-date-input')?.value || parsedRecapCache.date;
  const sessionTitle = parsedRecapCache.sessionTitle || 'Completed Workout';
  const exCount = parsedRecapCache.exercises.length;
  const exercisesSummary = exCount > 0 ? `${exCount} exercise${exCount === 1 ? '' : 's'}` : 'Workout Recap';

  if (!ath.logs) ath.logs = [];

  const newLog = {
    id: 'log-' + Date.now() + '-' + Math.random().toString(36).substr(2, 4),
    athleteId: ath.id,
    athleteName: ath.name,
    athleteInitials: ath.initials || ath.name.slice(0, 2).toUpperCase(),
    sessionId: 'imported',
    sessionTitle: sessionTitle,
    date: dateVal,
    status: 'completed',
    exercisesSummary: exercisesSummary,
    exercises: parsedRecapCache.exercises,
    notes: parsedRecapCache.notes || '',
    loggedAt: new Date().toISOString()
  };

  ath.logs.unshift(newLog);
  saveAthletesToStorage();

  closeModal('import-recap-modal');
  showToast(`Workout recap imported for ${ath.name}!`);

  if (currentView === 'calendar') {
    renderCalendarView();
  } else if (currentView === 'athlete') {
    renderAthleteDetail();
  }
}

// -------------------------------------------------------------
// ATHLETE PROFILE ACTIVITY & HISTORY SECTION
// -------------------------------------------------------------

function renderAthleteHistorySection(ath) {
  const container = document.getElementById('athlete-history-timeline');
  if (!container || !ath) return;

  const logs = ath.logs || [];
  if (logs.length === 0) {
    container.innerHTML = `
      <div style="text-align:center; padding:30px 16px; color:var(--text-dim); background:rgba(0,0,0,0.2); border-radius:10px; border:1px dashed var(--border); font-size:0.85em;">
        No workout sessions logged yet for ${escapeHtml(ath.name)}.<br>
        Use <strong>Log Workout</strong> or <strong>Import Recap</strong> to record training history.
      </div>
    `;
    return;
  }

  // Sort logs by date descending
  const sorted = [...logs].sort((a, b) => (b.date || '').localeCompare(a.date || ''));

  container.innerHTML = sorted.map(l => {
    const statusClass = l.status || 'completed';
    const exercisesDetails = (l.exercises && l.exercises.length > 0)
      ? l.exercises.map(e => `<span class="history-item-summary" style="display:inline-block; margin-right:8px;">${escapeHtml(e.exercise)}: <strong style="color:var(--teal);">${escapeHtml(e.sets)}</strong></span>`).join(' ')
      : `<span class="history-item-summary">${escapeHtml(l.exercisesSummary || 'Workout session')}</span>`;

    return `
      <div class="history-item-card">
        <div class="history-item-header">
          <div class="history-header-left">
            <span class="history-date-badge">${escapeHtml(l.date || 'No date')}</span>
            <span class="history-item-title">${escapeHtml(l.sessionTitle || 'Workout Session')}</span>
            <span class="cal-session-badge ${statusClass}" style="padding:2px 7px; font-size:0.7em;">${statusClass.toUpperCase()}</span>
          </div>
          <button type="button" class="history-btn-del" onclick="deleteCalendarLog('${l.id}', '${ath.id}')" title="Delete Log">
            ${MUI.delete}
          </button>
        </div>
        <div class="history-item-body">
          <div class="history-exercises-wrap">${exercisesDetails}</div>
          ${l.notes ? `<div class="history-item-notes">Coach Notes: "${escapeHtml(l.notes)}"</div>` : ''}
        </div>
      </div>
    `;
  }).join('');
}

// ==========================================
// 24/7 PRIVATE CLOUD SYNC ENGINE (GitHub Gist API)
// ==========================================
const CLOUD_SYNC_STORAGE_KEY = 'coach_studio_cloud_sync_cfg';
let cloudAutoPushTimer = null;

function getCloudSyncConfig() {
  try {
    const raw = localStorage.getItem(CLOUD_SYNC_STORAGE_KEY);
    if (raw) return JSON.parse(raw);
  } catch (e) {}
  return { token: '', gistId: '', autoSync: true, lastSync: null };
}

function saveCloudSyncConfig(cfg) {
  try {
    localStorage.setItem(CLOUD_SYNC_STORAGE_KEY, JSON.stringify(cfg));
  } catch (e) {}
}

function updateCloudSyncIndicator(state, message = '') {
  const dot = document.querySelector('.cloud-sync-status-dot');
  if (dot) {
    dot.className = 'cloud-sync-status-dot ' + (state || 'idle');
  }
  const statusEl = document.getElementById('cloud-sync-status-text');
  if (statusEl) {
    if (state === 'syncing') {
      statusEl.textContent = 'Syncing...';
      statusEl.style.color = 'var(--teal)';
    } else if (state === 'success') {
      statusEl.textContent = 'Connected & Synced';
      statusEl.style.color = '#10B981';
    } else if (state === 'error') {
      statusEl.textContent = message || 'Sync Error';
      statusEl.style.color = '#EF4444';
    } else {
      const cfg = getCloudSyncConfig();
      if (cfg.token) {
        statusEl.textContent = 'Ready (Idle)';
        statusEl.style.color = 'var(--text-main)';
      } else {
        statusEl.textContent = 'Not Configured';
        statusEl.style.color = 'var(--text-dim)';
      }
    }
  }
  const timeEl = document.getElementById('cloud-sync-last-time');
  if (timeEl) {
    const cfg = getCloudSyncConfig();
    if (cfg.lastSync) {
      const d = new Date(cfg.lastSync);
      timeEl.textContent = 'Last synced: ' + d.toLocaleDateString() + ' ' + d.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
    } else {
      timeEl.textContent = 'Never synced';
    }
  }
}

function openCloudSyncModal() {
  const cfg = getCloudSyncConfig();
  const tokenInput = document.getElementById('cloud-sync-token-input');
  const gistInput = document.getElementById('cloud-sync-gist-id-input');
  const autoCheckbox = document.getElementById('cloud-sync-auto-checkbox');

  if (tokenInput) tokenInput.value = cfg.token || '';
  if (gistInput) gistInput.value = cfg.gistId || '';
  if (autoCheckbox) autoCheckbox.checked = cfg.autoSync !== false;

  updateCloudSyncIndicator(cfg.token ? 'idle' : 'error');
  openModal('cloud-sync-modal');
}

function saveCloudSyncSettingsFromUI() {
  const tokenInput = document.getElementById('cloud-sync-token-input');
  const gistInput = document.getElementById('cloud-sync-gist-id-input');
  const autoCheckbox = document.getElementById('cloud-sync-auto-checkbox');

  const cfg = getCloudSyncConfig();
  cfg.token = (tokenInput ? tokenInput.value : '').trim();
  cfg.gistId = (gistInput ? gistInput.value : '').trim();
  cfg.autoSync = autoCheckbox ? autoCheckbox.checked : true;

  saveCloudSyncConfig(cfg);
  updateCloudSyncIndicator(cfg.token ? 'idle' : 'error');
  showToast('Cloud sync settings saved!');

  if (cfg.token && cfg.autoSync) {
    triggerManualCloudSync();
  }
}

function packStudioData() {
  return {
    version: 2,
    timestamp: new Date().toISOString(),
    lastModified: localStorage.getItem('coach_studio_last_modified') || new Date().toISOString(),
    athletes: studioAthletes,
    templates: sessionTemplates,
    activeAthleteId: activeAthleteId
  };
}

async function pushToCloudSync(isSilent = false) {
  const cfg = getCloudSyncConfig();
  if (!cfg.token) {
    if (!isSilent) showToast('GitHub Personal Access Token is required for Cloud Sync.');
    updateCloudSyncIndicator('error', 'Missing Token');
    return false;
  }

  updateCloudSyncIndicator('syncing');
  const payload = packStudioData();
  const gistContent = JSON.stringify(payload, null, 2);

  const reqBody = {
    description: "Coach Studio Private Backup (24/7 Mobile Sync)",
    files: {
      "coach_studio_data.json": {
        "content": gistContent
      }
    }
  };

  try {
    let url = 'https://api.github.com/gists';
    let method = 'POST';

    if (cfg.gistId) {
      url = `https://api.github.com/gists/${cfg.gistId}`;
      method = 'PATCH';
    } else {
      reqBody.public = false;
    }

    const res = await fetch(url, {
      method: method,
      headers: {
        'Authorization': `token ${cfg.token}`,
        'Accept': 'application/vnd.github.v3+json',
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(reqBody)
    });

    if (!res.ok) {
      const errJson = await res.json().catch(() => ({}));
      throw new Error(errJson.message || `HTTP ${res.status}`);
    }

    const data = await res.json();
    if (!cfg.gistId && data.id) {
      cfg.gistId = data.id;
      const gistInput = document.getElementById('cloud-sync-gist-id-input');
      if (gistInput) gistInput.value = data.id;
    }
    cfg.lastSync = new Date().toISOString();
    saveCloudSyncConfig(cfg);
    updateCloudSyncIndicator('success');
    if (!isSilent) showToast('Cloud sync upload complete!');
    return true;
  } catch (err) {
    console.error('Cloud push failed:', err);
    updateCloudSyncIndicator('error', err.message);
    if (!isSilent) showToast(`Cloud push error: ${err.message}`);
    return false;
  }
}

async function pullFromCloudSync(isSilent = false) {
  const cfg = getCloudSyncConfig();
  if (!cfg.token || !cfg.gistId) {
    if (!isSilent) showToast('Token and Gist ID required to pull cloud data.');
    updateCloudSyncIndicator('error', 'Not Configured');
    return false;
  }

  updateCloudSyncIndicator('syncing');
  try {
    const res = await fetch(`https://api.github.com/gists/${cfg.gistId}`, {
      method: 'GET',
      headers: {
        'Authorization': `token ${cfg.token}`,
        'Accept': 'application/vnd.github.v3+json'
      }
    });

    if (!res.ok) {
      const errJson = await res.json().catch(() => ({}));
      throw new Error(errJson.message || `HTTP ${res.status}`);
    }

    const data = await res.json();
    const file = data.files && data.files['coach_studio_data.json'];
    if (!file || !file.content) {
      throw new Error('Gist does not contain coach_studio_data.json');
    }

    const remoteData = JSON.parse(file.content);
    if (!Array.isArray(remoteData.athletes)) {
      throw new Error('Invalid format: athletes array missing');
    }

    studioAthletes = remoteData.athletes;
    if (Array.isArray(remoteData.templates) && remoteData.templates.length > 0) {
      sessionTemplates = remoteData.templates;
    }
    if (remoteData.activeAthleteId && studioAthletes.some(a => a.id === remoteData.activeAthleteId)) {
      activeAthleteId = remoteData.activeAthleteId;
    } else if (studioAthletes.length > 0) {
      activeAthleteId = studioAthletes[0].id;
    }

    localStorage.setItem('coach_studio_athletes', JSON.stringify(studioAthletes));
    localStorage.setItem('coach_studio_templates', JSON.stringify(sessionTemplates));
    if (remoteData.lastModified) {
      localStorage.setItem('coach_studio_last_modified', remoteData.lastModified);
    }

    cfg.lastSync = new Date().toISOString();
    saveCloudSyncConfig(cfg);
    updateCloudSyncIndicator('success');

    // Refresh UI
    renderHub();
    renderAthleteList();
    renderActiveAthlete();
    renderCalendar();

    if (!isSilent) showToast(`Cloud sync complete! Restored ${studioAthletes.length} athletes.`);
    return true;
  } catch (err) {
    console.error('Cloud pull failed:', err);
    updateCloudSyncIndicator('error', err.message);
    if (!isSilent) showToast(`Cloud pull error: ${err.message}`);
    return false;
  }
}

async function triggerManualCloudSync() {
  const cfg = getCloudSyncConfig();
  if (!cfg.token) {
    openCloudSyncModal();
    return;
  }

  if (!cfg.gistId) {
    await pushToCloudSync(false);
    return;
  }

  updateCloudSyncIndicator('syncing');
  try {
    const res = await fetch(`https://api.github.com/gists/${cfg.gistId}`, {
      method: 'GET',
      headers: {
        'Authorization': `token ${cfg.token}`,
        'Accept': 'application/vnd.github.v3+json'
      }
    });
    if (!res.ok) {
      await pushToCloudSync(false);
      return;
    }
    const gist = await res.json();
    const file = gist.files && gist.files['coach_studio_data.json'];
    if (!file || !file.content) {
      await pushToCloudSync(false);
      return;
    }
    const remoteData = JSON.parse(file.content);
    const remoteTime = new Date(remoteData.lastModified || gist.updated_at).getTime();
    const localTime = new Date(localStorage.getItem('coach_studio_last_modified') || 0).getTime();

    if (remoteTime > localTime) {
      await pullFromCloudSync(false);
    } else {
      await pushToCloudSync(false);
    }
  } catch (e) {
    console.warn('Sync comparison fallback to push:', e);
    await pushToCloudSync(false);
  }
}

function forcePushToCloud() {
  pushToCloudSync(false);
}

function forcePullFromCloud() {
  if (confirm('Pulling from cloud will overwrite local data on this device with remote cloud data. Proceed?')) {
    pullFromCloudSync(false);
  }
}

function scheduleCloudAutoPush() {
  const cfg = getCloudSyncConfig();
  if (!cfg.token || cfg.autoSync === false) return;
  if (cloudAutoPushTimer) clearTimeout(cloudAutoPushTimer);
  cloudAutoPushTimer = setTimeout(() => {
    pushToCloudSync(true);
  }, 2500);
}

async function initCloudSync() {
  const cfg = getCloudSyncConfig();
  updateCloudSyncIndicator(cfg.token ? 'idle' : 'error');
  if (cfg.token && cfg.gistId && cfg.autoSync !== false) {
    try {
      const res = await fetch(`https://api.github.com/gists/${cfg.gistId}`, {
        method: 'GET',
        headers: {
          'Authorization': `token ${cfg.token}`,
          'Accept': 'application/vnd.github.v3+json'
        }
      });
      if (res.ok) {
        const gist = await res.json();
        const file = gist.files && gist.files['coach_studio_data.json'];
        if (file && file.content) {
          const remoteData = JSON.parse(file.content);
          const remoteTime = new Date(remoteData.lastModified || gist.updated_at).getTime();
          const localTime = new Date(localStorage.getItem('coach_studio_last_modified') || 0).getTime();
          if (remoteTime > localTime) {
            await pullFromCloudSync(true);
          }
        }
      }
    } catch (e) {
      console.warn('Initial silent cloud sync check skipped:', e);
    }
  }
}

// ==========================================
// MASTER PASSCODE SECURITY ENGINE
// ==========================================
const PASSCODE_STORAGE_KEY = 'coach_studio_pin_config';
let enteredPinBuffer = '';

function getPasscodeConfig() {
  try {
    const raw = localStorage.getItem(PASSCODE_STORAGE_KEY);
    if (raw) return JSON.parse(raw);
  } catch (e) {}
  return { enabled: false, hash: '', salt: '' };
}

function savePasscodeConfig(cfg) {
  try {
    localStorage.setItem(PASSCODE_STORAGE_KEY, JSON.stringify(cfg));
  } catch (e) {}
}

async function hashPasscode(pin, salt) {
  const enc = new TextEncoder();
  const data = enc.encode(salt + ':' + pin + ':coach_studio_auth');
  const hashBuf = await crypto.subtle.digest('SHA-256', data);
  return Array.from(new Uint8Array(hashBuf)).map(b => b.toString(16).padStart(2, '0')).join('');
}

function openPasscodeSettingsModal() {
  const cfg = getPasscodeConfig();
  const newPinInput = document.getElementById('new-pin-input');
  const confirmPinInput = document.getElementById('confirm-pin-input');
  const disableBtn = document.getElementById('btn-disable-pin');
  const statusMsg = document.getElementById('passcode-status-msg');

  if (newPinInput) newPinInput.value = '';
  if (confirmPinInput) confirmPinInput.value = '';

  if (disableBtn) {
    disableBtn.style.display = cfg.enabled ? 'block' : 'none';
  }
  if (statusMsg) {
    statusMsg.innerHTML = cfg.enabled
      ? '<span style="color:var(--teal); font-weight:700;">PIN protection is currently ENABLED.</span> Enter a new PIN above to change it.'
      : '<span style="color:var(--text-dim);">PIN protection is currently disabled.</span>';
  }
  openModal('passcode-modal');
}

async function saveNewPinSecurity() {
  const newPinInput = document.getElementById('new-pin-input');
  const confirmPinInput = document.getElementById('confirm-pin-input');

  const p1 = (newPinInput ? newPinInput.value : '').trim();
  const p2 = (confirmPinInput ? confirmPinInput.value : '').trim();

  if (!/^\\d{4,6}$/.test(p1)) {
    showToast('PIN must be 4 to 6 digits (numbers only).');
    return;
  }
  if (p1 !== p2) {
    showToast('PIN entries do not match. Please re-enter.');
    return;
  }

  const salt = Math.random().toString(36).substring(2, 12);
  const hash = await hashPasscode(p1, salt);
  savePasscodeConfig({ enabled: true, hash: hash, salt: salt });

  closeModal('passcode-modal');
  showToast('PIN security saved! Coach Studio is protected.');
}

function disablePinSecurity() {
  if (confirm('Are you sure you want to disable PIN lock protection?')) {
    savePasscodeConfig({ enabled: false, hash: '', salt: '' });
    closeModal('passcode-modal');
    showToast('PIN lock disabled.');
  }
}

function lockStudioNow() {
  closeAllCustomSelects();
  const lockScreen = document.getElementById('passcode-lock-screen');
  if (lockScreen) {
    lockScreen.style.display = 'flex';
    clearPinInput();
    const hiddenInput = document.getElementById('pin-keyboard-input');
    if (hiddenInput) {
      hiddenInput.value = '';
      hiddenInput.focus();
    }
  }
}

function updatePinDotsUI() {
  const len = enteredPinBuffer.length;
  for (let i = 0; i < 6; i++) {
    const dot = document.getElementById('pdot-' + i);
    if (dot) {
      if (i < len) {
        dot.classList.add('filled');
      } else {
        dot.classList.remove('filled');
      }
    }
  }
}

function pressPinDigit(digit) {
  if (enteredPinBuffer.length < 6) {
    enteredPinBuffer += String(digit);
    updatePinDotsUI();
    const errEl = document.getElementById('pin-error-text');
    if (errEl) errEl.textContent = '';
    if (enteredPinBuffer.length >= 4) {
      checkEnteredPin();
    }
  }
}

function deletePinDigit() {
  if (enteredPinBuffer.length > 0) {
    enteredPinBuffer = enteredPinBuffer.slice(0, -1);
    updatePinDotsUI();
    const errEl = document.getElementById('pin-error-text');
    if (errEl) errEl.textContent = '';
  }
}

function clearPinInput() {
  enteredPinBuffer = '';
  updatePinDotsUI();
  const errEl = document.getElementById('pin-error-text');
  if (errEl) errEl.textContent = '';
}

function handlePinKeyInput(val) {
  const digitsOnly = val.replace(/\\D/g, '').slice(0, 6);
  enteredPinBuffer = digitsOnly;
  updatePinDotsUI();
  if (enteredPinBuffer.length >= 4) {
    checkEnteredPin();
  }
}

async function checkEnteredPin() {
  const cfg = getPasscodeConfig();
  if (!cfg.enabled) {
    const lockScreen = document.getElementById('passcode-lock-screen');
    if (lockScreen) lockScreen.style.display = 'none';
    return;
  }

  const enteredHash = await hashPasscode(enteredPinBuffer, cfg.salt);
  if (enteredHash === cfg.hash) {
    const lockScreen = document.getElementById('passcode-lock-screen');
    if (lockScreen) lockScreen.style.display = 'none';
    clearPinInput();
    showToast('Studio Unlocked. Welcome back!');
  } else if (enteredPinBuffer.length >= 6) {
    const dotsContainer = document.getElementById('pin-dots-container');
    if (dotsContainer) {
      dotsContainer.classList.add('pin-shake');
      setTimeout(() => dotsContainer.classList.remove('pin-shake'), 500);
    }
    const errEl = document.getElementById('pin-error-text');
    if (errEl) errEl.textContent = 'Incorrect PIN. Please try again.';
    setTimeout(() => {
      clearPinInput();
    }, 450);
  }
}

function initPasscodeSecurity() {
  const cfg = getPasscodeConfig();
  if (cfg.enabled) {
    lockStudioNow();
  }
}

"""

# Assemble builder.html safely without f-string brace escapes
builder_html_parts = [
    '<!DOCTYPE html>\n<html lang="en">\n<head>\n<meta charset="UTF-8">\n<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">\n<meta name="apple-mobile-web-app-capable" content="yes">\n<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">\n<meta name="mobile-web-app-capable" content="yes">\n<meta name="theme-color" content="#0d1117">\n<link rel="manifest" href="data:application/manifest+json,%7B%22name%22%3A%22Coach%20Studio%22%2C%22short_name%22%3A%22Coach%20Studio%22%2C%22start_url%22%3A%22.%22%2C%22display%22%3A%22standalone%22%2C%22background_color%22%3A%22%230d1117%22%2C%22theme_color%22%3A%22%230d1117%22%7D">\n<title>Coach Studio - Athlete Management</title>\n<style>\n',
    STUDIO_CSS,
    '\n</style>\n</head>\n<body>\n',
    STUDIO_HTML.format(**MUI_ICONS) if '{MUI_ICONS[' not in STUDIO_HTML else re.sub(r"\{MUI_ICONS\['(\w+)'\]\}", lambda m: MUI_ICONS.get(m.group(1), ''), STUDIO_HTML),
    '\n<script>\nconst MUI = ',
    mui_icons_json,
    ';\nwindow.INITIAL_ATHLETES = ',
    initial_athletes_json,
    ';\nconst CLIENT_CSS_TEMPLATE = ',
    client_css_json,
    ';\nconst CLIENT_JS_TEMPLATE = ',
    client_js_json,
    ';\n',
    STUDIO_JS,
    "\n// Initial Boot\ndocument.addEventListener('DOMContentLoaded', function() {\n  navigateTo('hub');\n  initCloudSync();\n  initPasscodeSecurity();\n});\n</script>\n</body>\n</html>"
]

builder_html_content = "".join(builder_html_parts)

output_path = PROJECT_DIR / 'builder.html'
with open(output_path, 'w', encoding='utf-8') as f:
    f.write(builder_html_content)

index_path = PROJECT_DIR / 'index.html'
with open(index_path, 'w', encoding='utf-8') as f:
    f.write(builder_html_content)

print(f"SUCCESS: Generated Multi-Athlete Coach Studio -> {output_path} & {index_path} ({len(builder_html_content)} bytes)")