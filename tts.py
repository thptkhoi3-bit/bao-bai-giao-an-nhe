"""Chuyển văn bản thành giọng nói (Gemini TTS) — dùng chung GEMINI_API_KEY qua module gemini.
3 chế độ: đọc thường (giọng có sẵn hoặc mô tả tự do), nghe–viết chính tả (ngắt nghỉ + lặp),
đối thoại (mỗi nhân vật một giọng). Đầu ra: WAV 24kHz 16-bit mono (nối PCM trực tiếp, không cần ffmpeg).
KHÔNG ghi key vào log."""
import base64, io, json, logging, re, urllib.request, urllib.error, wave

import gemini

log = logging.getLogger('tts')

_SILENCE_MP3 = base64.b64decode('''//OEwAAAAAAAAAAAAEluZm8AAAAPAAAALAAAGYAADQ0SEhgYHh4eIyMpKS8vNDQ0OjpAQEVFS0tLUFBWVlxcXGFhZ2dtbXJycnh4fn6Dg4mJiY+PlJSampqgoKWlq6uwsLC2try8wcHHx8fNzdLS2NjY3t7j4+np7+/v9PT6+v//AAAAAExhdmM2Mi4yOAAAAAAAAAAAAAAAACQDAAAAAAAAABmACaAqPgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA//NkxAAAAANIAAAAAExBTUVVVVVMQU1FMy4xMDBVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV//NkxHwAAANIAAAAAFVVVVVVVVVMQU1FMy4xMDBVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV//NkxHwAAANIAAAAAFVVVVVVVVVMQU1FMy4xMDBVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV//NkxHwAAANIAAAAAFVVVVVVVVVMQU1FMy4xMDBVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV//NkxHwAAANIAAAAAFVVVVVVVVVMQU1FMy4xMDBVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV//NkxHwAAANIAAAAAFVVVVVVVVVMQU1FMy4xMDBVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV//NkxHwAAANIAAAAAFVVVVVVVVVMQU1FMy4xMDBVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV//NkxHwAAANIAAAAAFVVVVVVVVVMQU1FMy4xMDBVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV//NkxHwAAANIAAAAAFVVVVVVVVVMQU1FMy4xMDBVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV//NkxHwAAANIAAAAAFVVVVVVVVVMQU1FMy4xMDBVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV//NkxHwAAANIAAAAAFVVVVVVVVVMQU1FMy4xMDBVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV//NkxHwAAANIAAAAAFVVVVVVVVVMQU1FMy4xMDBVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV//NkxHwAAANIAAAAAFVVVVVVVVVMQU1FMy4xMDBVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV//NkxHwAAANIAAAAAFVVVVVVVVVMQU1FMy4xMDBVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV//NkxHwAAANIAAAAAFVVVVVVVVVMQU1FMy4xMDBVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV//NkxHwAAANIAAAAAFVVVVVVVVVMQU1FMy4xMDBVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV//NkxHwAAANIAAAAAFVVVVVVVVVMQU1FMy4xMDBVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV//NkxHwAAANIAAAAAFVVVVVVVVVMQU1FMy4xMDBVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV//NkxHwAAANIAAAAAFVVVVVVVVVMQU1FMy4xMDBVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV//NkxHwAAANIAAAAAFVVVVVVVVVMQU1FMy4xMDBVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV//NkxHwAAANIAAAAAFVVVVVVVVVMQU1FMy4xMDBVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV//NkxHwAAANIAAAAAFVVVVVVVVVMQU1FMy4xMDBVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV//NkxHwAAANIAAAAAFVVVVVVVVVMQU1FMy4xMDBVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV//NkxHwAAANIAAAAAFVVVVVVVVVMQU1FMy4xMDBVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV//NkxHwAAANIAAAAAFVVVVVVVVVMQU1FMy4xMDBVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV//NkxHwAAANIAAAAAFVVVVVVVVVMQU1FMy4xMDBVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV//NkxHwAAANIAAAAAFVVVVVVVVVMQU1FMy4xMDBVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV//NkxHwAAANIAAAAAFVVVVVVVVVMQU1FMy4xMDBVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV//NkxHwAAANIAAAAAFVVVVVVVVVMQU1FMy4xMDBVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV//NkxHwAAANIAAAAAFVVVVVVVVVMQU1FMy4xMDBVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV//NkxHwAAANIAAAAAFVVVVVVVVVMQU1FMy4xMDBVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV//NkxHwAAANIAAAAAFVVVVVVVVVMQU1FMy4xMDBVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV//NkxHwAAANIAAAAAFVVVVVVVVVMQU1FMy4xMDBVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV//NkxHwAAANIAAAAAFVVVVVVVVVMQU1FMy4xMDBVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV//NkxHwAAANIAAAAAFVVVVVVVVVMQU1FMy4xMDBVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV//NkxHwAAANIAAAAAFVVVVVVVVVMQU1FMy4xMDBVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV//NkxHwAAANIAAAAAFVVVVVVVVVMQU1FMy4xMDBVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV//NkxHwAAANIAAAAAFVVVVVVVVVMQU1FMy4xMDBVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV//NkxHwAAANIAAAAAFVVVVVVVVVMQU1FMy4xMDBVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV//NkxHwAAANIAAAAAFVVVVVVVVVMQU1FMy4xMDBVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV//NkxHwAAANIAAAAAFVVVVVVVVVMQU1FMy4xMDBVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV//NkxHwAAANIAAAAAFVVVVVVVVVMQU1FMy4xMDBVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV//NkxHwAAANIAAAAAFVVVVVVVVVMQU1FMy4xMDBVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV//NkxHwAAANIAAAAAFVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV''')  # MP3 im lang 1 giay 24kHz mono 48kbps, lap lai theo so giay can

TTS_MODELS = ['gemini-2.5-flash-preview-tts']
RATE = 24000
MAX_CHARS = 3000          # giới hạn ký tự / 1 lần tạo, chống cháy quota
CHUNK = 600               # mỗi lần gọi AI tối đa ~600 ký tự (gộp theo câu)
PAUSE_TURN = 0.6          # giây nghỉ giữa các lượt thoại

VOICES = {
    'co_giao':   dict(name='Kore',    icon='👩', label='Cô giáo',   style='giọng nữ dịu dàng, ấm áp, đọc rõ ràng với tốc độ vừa phải'),
    'thay_giao': dict(name='Charon',  icon='👨', label='Thầy giáo', style='giọng nam trầm ấm, đọc rõ ràng với tốc độ vừa phải'),
    'be_gai':    dict(name='Leda',    icon='👧', label='Bé gái',    style='giọng bé gái Việt Nam khoảng 7 tuổi, trong trẻo, hồn nhiên, vui tươi'),
    'be_trai':   dict(name='Puck',    icon='👦', label='Bé trai',   style='giọng bé trai Việt Nam khoảng 7 tuổi, lanh lợi, trong sáng, vui vẻ'),
    'ba':        dict(name='Sulafat', icon='👵', label='Bà',        style='giọng bà lão hiền từ, chậm rãi, ấm áp'),
    'ong':       dict(name='Fenrir',  icon='🧓', label='Ông',       style='giọng ông lão trầm, chậm rãi, hiền hậu'),
}
SPEEDS = {'cham': ' Đọc chậm rãi, rõ từng chữ.', 'vua': '', 'nhanh': ' Đọc hơi nhanh, trôi chảy.'}


def voices_public():
    return [dict(id=k, icon=v['icon'], label=v['label']) for k, v in VOICES.items()]


def _resolve(voice_id, style_custom):
    """Trả về (voice_name, style). voice_id='custom' -> dùng mô tả tự do + giọng nền Leda."""
    if voice_id == 'custom':
        st = (style_custom or '').strip()[:300]
        return 'Leda', ('giọng đọc theo mô tả sau: ' + st) if st else VOICES['co_giao']['style']
    v = VOICES.get(voice_id) or VOICES['co_giao']
    return v['name'], v['style']


def _pcm(text, voice_name, style, speed='vua'):
    """Gọi Gemini TTS, trả về bytes PCM 24kHz 16-bit mono. Lỗi -> ValueError (thông báo thân thiện)."""
    prompt = f'Đọc đoạn văn sau bằng tiếng Việt với {style}.{SPEEDS.get(speed, "")}\nChỉ đọc đúng nội dung, không thêm bớt:\n{text}'
    body = {'contents': [{'parts': [{'text': prompt}]}],
            'generationConfig': {'responseModalities': ['AUDIO'],
                                 'speechConfig': {'voiceConfig': {'prebuiltVoiceConfig': {'voiceName': voice_name}}}}}
    k = gemini.key()
    last = None
    for m in TTS_MODELS:
        rq = urllib.request.Request(f'{gemini.base_url()}/models/{m}:generateContent',
                                    data=json.dumps(body).encode(), method='POST',
                                    headers={'Content-Type': 'application/json', 'x-goog-api-key': k})
        try:
            with urllib.request.urlopen(rq, timeout=150) as r:
                res = json.loads(r.read())
        except urllib.error.HTTPError as e:
            last = e
            if e.code == 404:
                continue
            if e.code == 429:
                raise ValueError('Dịch vụ giọng đọc đang quá tải, thầy/cô thử lại sau ít phút nhé ạ.')
            if e.code in (400, 401, 403):
                raise ValueError('Mã AI (GEMINI_API_KEY) không hợp lệ — thầy/cô báo quản trị nhé ạ.')
            raise ValueError('Chưa kết nối được dịch vụ giọng đọc, thầy/cô thử lại sau nhé ạ.')
        except Exception as e:
            log.warning('tts: lỗi kết nối %s', type(e).__name__)
            raise ValueError('Chưa kết nối được dịch vụ giọng đọc, thầy/cô thử lại sau nhé ạ.')
        try:
            for p in res['candidates'][0]['content']['parts']:
                d = (p.get('inlineData') or {}).get('data')
                if d:
                    return base64.b64decode(d)
        except Exception:
            pass
        raise ValueError('Dịch vụ giọng đọc chưa trả về âm thanh cho đoạn này, thầy/cô thử rút ngắn văn bản nhé ạ.')
    raise ValueError('Mô hình giọng đọc chưa khả dụng — thầy/cô báo quản trị nhé ạ.')


def split_sentences(text):
    parts = re.split(r'(?<=[.!?…])\s+', (text or '').replace('\r', ' ').strip())
    return [p.strip() for p in parts if p.strip()]


def chunk_text(text, limit=CHUNK):
    out, cur = [], ''
    for s in split_sentences(text):
        if len(s) > limit:  # câu quá dài: cắt cứng theo khoảng trắng
            words, piece = s.split(), ''
            pieces = []
            for w in words:
                if len(piece) + len(w) + 1 > limit and piece:
                    pieces.append(piece); piece = w
                else:
                    piece = (piece + ' ' + w).strip()
            if piece: pieces.append(piece)
            s_list = pieces
        else:
            s_list = [s]
        for s2 in s_list:
            if len(cur) + len(s2) + 1 > limit and cur:
                out.append(cur); cur = s2
            else:
                cur = (cur + ' ' + s2).strip()
    if cur: out.append(cur)
    return out


def _silence(sec):
    return b'\x00' * int(RATE * 2 * max(0.0, float(sec)))


def pcm_to_wav(pcm):
    buf = io.BytesIO()
    with wave.open(buf, 'wb') as w:
        w.setnchannels(1); w.setsampwidth(2); w.setframerate(RATE)
        w.writeframes(pcm)
    return buf.getvalue()


# ---------- Động cơ chính: Microsoft Edge TTS (giọng Việt tự nhiên, miễn phí) ----------
#               (giọng gốc, tốc độ %, tông Hz) — giọng trẻ em/lão tạo bằng nâng/hạ tông
PRESETS_EDGE = {
    'co_giao':   ('vi-VN-HoaiMyNeural', 0, 0),
    'thay_giao': ('vi-VN-NamMinhNeural', 0, 0),
    'be_gai':    ('vi-VN-HoaiMyNeural', 2, 48),
    'be_trai':   ('vi-VN-NamMinhNeural', 2, 32),
    'ba':        ('vi-VN-HoaiMyNeural', -14, -12),
    'ong':       ('vi-VN-NamMinhNeural', -12, -8),
}
SPEED_EDGE = {'cham': -20, 'vua': 0, 'nhanh': 15}


def _edge_mp3_once(text, voice, rate_pct, pitch_hz):
    import asyncio, edge_tts, tempfile, os
    async def go():
        c = edge_tts.Communicate(text, voice, rate=f'{rate_pct:+d}%', pitch=f'{pitch_hz:+d}Hz')
        with tempfile.NamedTemporaryFile(suffix='.mp3', delete=False) as f:
            path = f.name
        try:
            await c.save(path)
            with open(path, 'rb') as f: return f.read()
        finally:
            try: os.remove(path)
            except OSError: pass
    return asyncio.run(go())


def _edge_preset(voice_id, speed='vua'):
    v, r, p = PRESETS_EDGE.get(voice_id) or PRESETS_EDGE['co_giao']
    return v, r + SPEED_EDGE.get(speed, 0), p


def _edge_read(text, voice_id, speed='vua'):
    v, r, p = _edge_preset(voice_id, speed)
    data = b''.join(_edge_mp3_once(c, v, r, p) for c in chunk_text(text))
    return data, 'mp3', round(len(data) / 6000, 1)   # MP3 48kbit/s ≈ 6000 byte/s


def _gemini_read_wav(text, voice_id='co_giao', style_custom='', speed='vua'):
    text = (text or '').strip()
    if not text: raise ValueError('Thầy/cô nhập văn bản cần đọc nhé ạ.')
    if len(text) > MAX_CHARS: raise ValueError(f'Văn bản quá dài (tối đa {MAX_CHARS} ký tự mỗi lần tạo).')
    vn, style = _resolve(voice_id, style_custom)
    pcm = b''.join(_pcm(c, vn, style, speed) for c in chunk_text(text))
    return pcm_to_wav(pcm)


def _gemini_dictation_wav(text, voice_id='co_giao', pause=5, repeat=2):
    text = (text or '').strip()
    if not text: raise ValueError('Thầy/cô nhập văn bản cần đọc nhé ạ.')
    if len(text) > MAX_CHARS: raise ValueError(f'Văn bản quá dài (tối đa {MAX_CHARS} ký tự mỗi lần tạo).')
    vn, style = _resolve(voice_id, '')
    pause = min(max(int(pause or 5), 1), 15); repeat = min(max(int(repeat or 2), 1), 3)
    parts = []
    for s in split_sentences(text):
        one = _pcm(s, vn, style, 'cham')
        for i in range(repeat):
            parts.append(one)
            parts.append(_silence(pause if i < repeat - 1 else pause + 2))
    return pcm_to_wav(b''.join(parts))


LINE_RE = re.compile(r'^\s*([^:\n]{1,24}?)\s*:\s*(.+)$')


def parse_dialogue(script):
    """Trả về [(nhân vật hoặc '' nếu dẫn chuyện, câu thoại)] theo từng dòng."""
    out = []
    for line in (script or '').splitlines():
        line = line.strip()
        if not line: continue
        m = LINE_RE.match(line)
        if m and not m.group(1).strip().isdigit():
            out.append((m.group(1).strip(), m.group(2).strip()))
        elif out and out[-1][0] == '__cont__':
            out[-1] = ('__cont__', out[-1][1] + ' ' + line)
        else:
            out.append(('', line))
    # gộp các đoạn dẫn chuyện liên tiếp? giữ nguyên từng dòng để giữ ngắt nghỉ tự nhiên
    return [(('' if n == '__cont__' else n), t) for n, t in out]


def dialogue_speakers(script):
    seen = []
    for name, _ in parse_dialogue(script):
        if name and name not in seen: seen.append(name)
    return seen


def _gemini_dialogue_wav(script, cast=None, narrator_voice='co_giao'):
    cast = cast or {}
    lines = parse_dialogue(script)
    if not lines: raise ValueError('Thầy/cô nhập kịch bản đối thoại nhé ạ.')
    total = sum(len(t) for _, t in lines)
    if total > MAX_CHARS: raise ValueError(f'Kịch bản quá dài (tối đa {MAX_CHARS} ký tự mỗi lần tạo).')
    parts = []
    for name, text in lines:
        vid = cast.get(name) if name else narrator_voice
        vn, style = _resolve(vid or 'co_giao', '')
        parts.append(_pcm(text, vn, style))
        parts.append(_silence(PAUSE_TURN))
    return pcm_to_wav(b''.join(parts))


# ---------- Hàm chính (endpoint gọi): Edge trước, Gemini dự phòng ----------
def synth_read(text, voice_id='co_giao', style_custom='', speed='vua'):
    """Trả về (data, ext, seconds). 'custom' -> Gemini diễn theo mô tả."""
    text = (text or '').strip()
    if not text: raise ValueError('Thầy/cô nhập văn bản cần đọc nhé ạ.')
    if len(text) > MAX_CHARS: raise ValueError(f'Văn bản quá dài (tối đa {MAX_CHARS} ký tự mỗi lần tạo).')
    if voice_id == 'custom':
        wav = _gemini_read_wav(text, 'custom', style_custom, speed)
        return wav, 'wav', round((len(wav) - 44) / 48000, 1)
    try:
        return _edge_read(text, voice_id, speed)
    except Exception as e:
        log.warning('tts: edge loi (%s) – dung gemini', type(e).__name__)
    wav = _gemini_read_wav(text, voice_id, '', speed)
    return wav, 'wav', round((len(wav) - 44) / 48000, 1)


def synth_dictation(text, voice_id='co_giao', pause=5, repeat=2):
    text = (text or '').strip()
    if not text: raise ValueError('Thầy/cô nhập văn bản cần đọc nhé ạ.')
    if len(text) > MAX_CHARS: raise ValueError(f'Văn bản quá dài (tối đa {MAX_CHARS} ký tự mỗi lần tạo).')
    pause = min(max(int(pause or 5), 1), 15); repeat = min(max(int(repeat or 2), 1), 3)
    if voice_id in PRESETS_EDGE:
        try:
            v, r, p = _edge_preset(voice_id, 'cham')
            parts = []
            for s in split_sentences(text):
                one = _edge_mp3_once(s, v, r, p)
                for i in range(repeat):
                    parts.append(one)
                    parts.append(_SILENCE_MP3 * (pause if i < repeat - 1 else pause + 2))
            data = b''.join(parts)
            return data, 'mp3', round(len(data) / 6000, 1)
        except Exception as e:
            log.warning('tts: edge loi (%s) – dung gemini', type(e).__name__)
    wav = _gemini_dictation_wav(text, voice_id if voice_id in VOICES else 'co_giao', pause, repeat)
    return wav, 'wav', round((len(wav) - 44) / 48000, 1)


def synth_dialogue(script, cast=None, narrator_voice='co_giao'):
    cast = cast or {}
    lines = parse_dialogue(script)
    if not lines: raise ValueError('Thầy/cô nhập kịch bản đối thoại nhé ạ.')
    total = sum(len(t) for _, t in lines)
    if total > MAX_CHARS: raise ValueError(f'Kịch bản quá dài (tối đa {MAX_CHARS} ký tự mỗi lần tạo).')
    try:
        parts = []
        for name, text in lines:
            vid = (cast.get(name) if name else narrator_voice) or 'co_giao'
            v, r, p = _edge_preset(vid, 'vua')
            parts.append(_edge_mp3_once(text, v, r, p))
            parts.append(_SILENCE_MP3)   # nghỉ 1 giây giữa các lượt thoại
        data = b''.join(parts)
        return data, 'mp3', round(len(data) / 6000, 1)
    except Exception as e:
        log.warning('tts: edge loi (%s) – dung gemini', type(e).__name__)
    wav = _gemini_dialogue_wav(script, cast, narrator_voice)
    return wav, 'wav', round((len(wav) - 44) / 48000, 1)
