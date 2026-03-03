import azure.cognitiveservices.speech as speechsdk
from pathlib import Path

# =====================================
# 1️⃣ CONFIGURATION (keep same keys/services)
# =====================================

AZURE_SPEECH_KEY = ""
AZURE_REGION = "westeurope"

# Output folder structure kept the same style as your script
BASE_OUTPUT_DIR = Path("/Users/palszirmai/Documents/AFI_003/")
SHEET_NAME = "HE"  # folder name only (no Excel reading)
OUTPUT_DIR = BASE_OUTPUT_DIR / SHEET_NAME
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# Text comes from code (not Excel)
HEBREW_TEXT = "זהו קטע דוגמה לקול עבור התערוכה ההונגרית"

# =====================================
# 2️⃣ Initialize Azure TTS
# =====================================

speech_config = speechsdk.SpeechConfig(
    subscription=AZURE_SPEECH_KEY,
    region=AZURE_REGION
)

# Output = WAV 48kHz 16-bit PCM (mono)
speech_config.set_speech_synthesis_output_format(
    speechsdk.SpeechSynthesisOutputFormat.Riff48Khz16BitMonoPcm
)

# =====================================
# 3️⃣ Voice Listing
# =====================================

def list_hebrew_voices():
    """
    Retrieve all available voices from this Azure Speech resource
    and return only Hebrew ones (locale starts with 'he-').
    """
    synthesizer = speechsdk.SpeechSynthesizer(speech_config=speech_config, audio_config=None)
    result = synthesizer.get_voices_async().get()

    if result.reason != speechsdk.ResultReason.VoicesListRetrieved:
        raise RuntimeError(f"Failed to retrieve voices: {result.reason}")

    hebrew = [v for v in result.voices if (v.locale or "").lower().startswith("he-")]
    hebrew.sort(key=lambda v: (v.locale, v.short_name))
    return hebrew

# =====================================
# 4️⃣ Speech Generation
# =====================================

def synthesize_audio(voice_name: str, text: str, output_path: Path):
    """Generate WAV using Azure Neural TTS."""
    try:
        # Skip if file exists
        if output_path.exists():
            print(f"⏭️ Skipping (already exists): {output_path.name}")
            return

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
            # cancellation_details is the reliable way to get error info
            try:
                cancellation = speechsdk.SpeechSynthesisCancellationDetails.from_result(result)
                if cancellation and cancellation.error_details:
                    print("Details:", cancellation.error_details)
            except Exception:
                pass

    except Exception as e:
        print(f"❌ Exception: {e}")

# =====================================
# 5️⃣ RUN
# =====================================

if __name__ == "__main__":
    print(f"📂 Output folder: {OUTPUT_DIR}\n")

    # 1) List Hebrew voices
    voices = list_hebrew_voices()

    print(f"Found {len(voices)} Hebrew voices:\n")
    for v in voices:
        print(f"- {v.short_name} | {v.gender} | {v.locale} | {v.voice_type}")

    # 2) Generate one sample per voice
    print("\n🔊 Generating samples...\n")
    for v in voices:
        # Keep output naming predictable and filesystem-safe
        filename = f"{v.short_name}_sample.wav"
        output_path = OUTPUT_DIR / filename
        synthesize_audio(v.short_name, HEBREW_TEXT, output_path)

    print("\n🎉 Done!")
