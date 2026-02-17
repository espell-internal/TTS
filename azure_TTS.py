import azure.cognitiveservices.speech as speechsdk
import pandas as pd
from pathlib import Path
import os

# =====================================
# 1️⃣ CONFIGURATION
# =====================================

# Azure Keys (replace with espell's)
AZURE_SPEECH_KEY = ""
AZURE_REGION = "westeurope"

# Excel source file — your exact path
XLSX_FILE = Path("/Users/palszirmai/Documents/AFI_003/AFI_003_text_prepped.xlsx")

# 🔧 EDIT THIS LINE TO CHOOSE WHICH SHEET TO READ
SHEET_NAME = "ENG_korr"   # ← change this to any sheet name

# Base output folder
BASE_OUTPUT_DIR = Path("/Users/palszirmai/Documents/AFI_003/")

# Output folder = 
OUTPUT_DIR = BASE_OUTPUT_DIR / SHEET_NAME
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


# =====================================
# 2️⃣ Initialize Azure TTS
# =====================================

speech_config = speechsdk.SpeechConfig(
    subscription=AZURE_SPEECH_KEY,
    region=AZURE_REGION
)

# Output = highest Azure quality: WAV 48kHz 16-bit PCM
speech_config.set_speech_synthesis_output_format(
    speechsdk.SpeechSynthesisOutputFormat.Riff48Khz16BitMonoPcm
)


# =====================================
# 3️⃣ Speech Generation
# =====================================

def synthesize_audio(voice_name, text, output_path):
    """Generate WAV using Azure Neural TTS."""
    try:
        # Use the selected voice
        speech_config.speech_synthesis_voice_name = voice_name

        audio_output = speechsdk.AudioConfig(filename=str(output_path))
        synthesizer = speechsdk.SpeechSynthesizer(
            speech_config=speech_config,
            audio_config=audio_output
        )

        print(f"Generating {output_path.name}  (voice={voice_name})")
        result = synthesizer.speak_text_async(text).get()

        if result.reason == speechsdk.ResultReason.SynthesizingAudioCompleted:
            print(f"✔ Saved: {output_path}")
        else:
            print(f"❌ Azure TTS Error: {result.reason}")
            cancellation = result.cancellation_details
            if cancellation:
                print("Details:", cancellation.error_details)

    except Exception as e:
        print(f"❌ Exception: {e}")


# =====================================
# 4️⃣ Process Selected Sheet
# =====================================

def process_sheet(xlsx_file, sheet_name):
    df = pd.read_excel(xlsx_file, sheet_name=sheet_name)

    for idx, row in df.iterrows():
        filename = str(row.iloc[0]).strip()
        voice    = str(row.iloc[1]).strip()
        text     = str(row.iloc[2]).strip()


        # Ensure WAV extension
        if not filename.lower().endswith(".wav"):
            filename += ".wav"

        output_path = OUTPUT_DIR / filename

        # 🔍 Skip if file exists (CHECK FIRST)
        if output_path.exists():
            print(f"⏭️ Skipping (already exists): {filename}")
            continue

        # Generate file only if it doesn't exist
        synthesize_audio(voice, text, output_path)


# =====================================
# 5️⃣ RUN
# =====================================

if __name__ == "__main__":
    print(f"📘 Processing sheet: {SHEET_NAME}")
    print(f"📄 Excel file: {XLSX_FILE}")
    print(f"📂 Output folder: {OUTPUT_DIR}\n")

    process_sheet(XLSX_FILE, SHEET_NAME)

    print("\n🎉 All audio files generated successfully!")
