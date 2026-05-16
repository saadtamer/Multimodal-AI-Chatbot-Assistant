import os
import base64
import requests
import streamlit as st
import PyPDF2
import pandas as pd
from docx import Document
from dotenv import load_dotenv

load_dotenv()

APP_TITLE = "Saad AI Chatbot"
APP_SUBTITLE = "Multi-task AI assistant powered by OpenRouter"
APP_DEVELOPER = "Eng. Saad Tamer Abo-Elazm"

MAX_PDF_PAGES = 8
MAX_DOCX_PARAGRAPHS = 250
MAX_TEXT_CHARS = 12000
MAX_EXCEL_ROWS = 300
MAX_HISTORY_MESSAGES = 12

AUDIO_INPUT_MODELS = {
    "nvidia/nemotron-3-nano-omni-30b-a3b-reasoning:free",
}

TASK_MODELS = {
    "🤖 محادثة عامة": {
        "default": "nvidia/nemotron-3-super-120b-a12b:free",
        "options": {
            "nvidia/nemotron-3-super-120b-a12b:free": "Nemotron 3 Super 120B",
            "openai/gpt-oss-120b:free": "GPT OSS 120B",
            "openrouter/owl-alpha": "Owl Alpha",
            "openai/gpt-oss-20b:free": "GPT OSS 20B",
            "z-ai/glm-4.5-air:free": "GLM 4.5 Air",
        },
    },
    "💻 كود": {
        "default": "poolside/laguna-m.1:free",
        "options": {
            "poolside/laguna-m.1:free": "Laguna M.1",
            "poolside/laguna-xs.2:free": "Laguna XS.2",
            "inclusionai/ring-2.6-1t:free": "Ring 2.6 1T",
            "nvidia/nemotron-3-super-120b-a12b:free": "Nemotron 3 Super 120B",
            "baidu/cobuddy:free": "CoBuddy",
            "openai/gpt-oss-120b:free": "GPT OSS 120B",
        },
    },
    "📝 تلخيص": {
        "default": "minimax/minimax-m2.5:free",
        "options": {
            "minimax/minimax-m2.5:free": "MiniMax M2.5",
            "openrouter/owl-alpha": "Owl Alpha",
            "nvidia/nemotron-3-nano-30b-a3b:free": "Nemotron Nano 30B",
            "openai/gpt-oss-120b:free": "GPT OSS 120B",
        },
    },
    "🌐 ترجمة متخصصة": {
        "default": "z-ai/glm-4.5-air:free",
        "options": {
            "z-ai/glm-4.5-air:free": "GLM 4.5 Air",
            "minimax/minimax-m2.5:free": "MiniMax M2.5",
            "nvidia/nemotron-3-nano-30b-a3b:free": "Nemotron Nano 30B",
            "openrouter/owl-alpha": "Owl Alpha",
        },
    },
    "🖼️ تحليل صورة": {
        "default": "google/gemma-4-31b-it:free",
        "options": {
            "google/gemma-4-31b-it:free": "Gemma 4 31B",
            "google/gemma-4-26b-a4b-it:free": "Gemma 4 26B",
            "nvidia/nemotron-nano-12b-v2-vl:free": "Nemotron Nano 12B VL",
            "nvidia/nemotron-3-nano-omni-30b-a3b-reasoning:free": "Nemotron Nano Omni",
        },
    },
    "🎤 صوت": {
        "default": "nvidia/nemotron-3-nano-omni-30b-a3b-reasoning:free",
        "options": {
            "nvidia/nemotron-3-nano-omni-30b-a3b-reasoning:free": "Nemotron Nano Omni",
        },
    },
    "📊 تحليل بيانات": {
        "default": "nvidia/nemotron-3-super-120b-a12b:free",
        "options": {
            "nvidia/nemotron-3-super-120b-a12b:free": "Nemotron 3 Super 120B",
            "openai/gpt-oss-120b:free": "GPT OSS 120B",
            "minimax/minimax-m2.5:free": "MiniMax M2.5",
            "openrouter/owl-alpha": "Owl Alpha",
            "inclusionai/ring-2.6-1t:free": "Ring 2.6 1T",
        },
    },
    "🎓 شرح وتعليم": {
        "default": "nvidia/nemotron-3-super-120b-a12b:free",
        "options": {
            "nvidia/nemotron-3-super-120b-a12b:free": "Nemotron 3 Super 120B",
            "openai/gpt-oss-120b:free": "GPT OSS 120B",
            "openrouter/owl-alpha": "Owl Alpha",
            "google/gemma-4-31b-it:free": "Gemma 4 31B",
            "nvidia/nemotron-nano-9b-v2:free": "Nemotron Nano 9B",
        },
    },
    "💼 CV وإيميلات": {
        "default": "nvidia/nemotron-3-super-120b-a12b:free",
        "options": {
            "nvidia/nemotron-3-super-120b-a12b:free": "Nemotron 3 Super 120B",
            "openai/gpt-oss-120b:free": "GPT OSS 120B",
            "minimax/minimax-m2.5:free": "MiniMax M2.5",
            "openrouter/owl-alpha": "Owl Alpha",
        },
    },
    "🔍 بحث وتلخيص مقالات": {
        "default": "openrouter/owl-alpha",
        "options": {
            "openrouter/owl-alpha": "Owl Alpha",
            "nvidia/nemotron-3-super-120b-a12b:free": "Nemotron 3 Super 120B",
            "openai/gpt-oss-120b:free": "GPT OSS 120B",
            "minimax/minimax-m2.5:free": "MiniMax M2.5",
        },
    },
    "🧮 مسائل رياضية": {
        "default": "inclusionai/ring-2.6-1t:free",
        "options": {
            "inclusionai/ring-2.6-1t:free": "Ring 2.6 1T",
            "nvidia/nemotron-3-super-120b-a12b:free": "Nemotron 3 Super 120B",
            "openai/gpt-oss-120b:free": "GPT OSS 120B",
            "nvidia/nemotron-nano-9b-v2:free": "Nemotron Nano 9B",
            "google/gemma-4-31b-it:free": "Gemma 4 31B",
        },
    },
    "🏥 استشارة طبية": {
        "default": "openai/gpt-oss-120b:free",
        "options": {
            "openai/gpt-oss-120b:free": "GPT OSS 120B",
            "openrouter/owl-alpha": "Owl Alpha",
            "nvidia/nemotron-3-super-120b-a12b:free": "Nemotron 3 Super 120B",
            "google/gemma-4-31b-it:free": "Gemma 4 31B",
        },
    },
    "⚖️ استشارة قانونية": {
        "default": "openrouter/owl-alpha",
        "options": {
            "openrouter/owl-alpha": "Owl Alpha",
            "openai/gpt-oss-120b:free": "GPT OSS 120B",
            "nvidia/nemotron-3-super-120b-a12b:free": "Nemotron 3 Super 120B",
            "nvidia/nemotron-3-nano-30b-a3b:free": "Nemotron Nano 30B",
        },
    },
}

TASK_DESCRIPTIONS = {
    "🤖 محادثة عامة": "محادثة طبيعية، أفكار، أسئلة عامة، وتنظيم معلومات.",
    "💻 كود": "كتابة كود، debugging، شرح ملفات، وتحسين المشاريع.",
    "📝 تلخيص": "تلخيص ملفات، محاضرات، مقالات، ونصوص طويلة.",
    "🌐 ترجمة متخصصة": "ترجمة احترافية مع الحفاظ على المصطلحات التقنية.",
    "🖼️ تحليل صورة": "تحليل صور واستخراج تفاصيل أو نصوص منها.",
    "🎤 صوت": "التعامل مع محتوى صوتي أو audio transcription.",
    "📊 تحليل بيانات": "تحليل CSV وExcel واستخراج insights.",
    "🎓 شرح وتعليم": "شرح خطوة بخطوة بأسلوب بسيط.",
    "💼 CV وإيميلات": "كتابة CV، emails، LinkedIn، وcover letters.",
    "🔍 بحث وتلخيص مقالات": "تلخيص وتحليل المقالات واستخراج أهم النقاط.",
    "🧮 مسائل رياضية": "حل مسائل رياضية بخطوات منظمة.",
    "🏥 استشارة طبية": "معلومات طبية عامة وليست بديلًا للطبيب.",
    "⚖️ استشارة قانونية": "معلومات قانونية عامة وليست بديلًا للمحامي.",
}

QUICK_PROMPTS = {
    "🤖 محادثة عامة": "اقترح لي خطة بسيطة أبدأ بها اليوم.",
    "💻 كود": "اشرح لي الكود خطوة بخطوة واقترح تحسينات.",
    "📝 تلخيص": "لخص المحتوى في نقاط واضحة ومنظمة.",
    "🌐 ترجمة متخصصة": "ترجم النص ترجمة احترافية مع الحفاظ على المصطلحات.",
    "🖼️ تحليل صورة": "حلل الصورة بالتفصيل واذكر كل الملاحظات.",
    "🎤 صوت": "استخرج الفكرة الأساسية من المحتوى الصوتي.",
    "📊 تحليل بيانات": "حلل البيانات واستخرج أهم insights.",
    "🎓 شرح وتعليم": "اشرح الموضوع كأني مبتدئ تمامًا.",
    "💼 CV وإيميلات": "اكتب لي email احترافي قصير ومناسب.",
    "🔍 بحث وتلخيص مقالات": "استخرج أهم الأفكار والنتائج من المقال.",
    "🧮 مسائل رياضية": "حل المسألة خطوة بخطوة.",
    "🏥 استشارة طبية": "اشرح لي المعلومة الطبية العامة بطريقة بسيطة.",
    "⚖️ استشارة قانونية": "اشرح لي المعلومة القانونية بشكل عام وبسيط.",
}


def model_supports_audio(model):
    return model in AUDIO_INPUT_MODELS


def get_api_key():
    manual_key = st.session_state.get("manual_api_key", "").strip()
    env_key = os.getenv("OPENROUTER_KEY") or os.getenv("OPENROUTER_API_KEY")
    return manual_key or env_key


def get_all_model_names():
    models = {}
    for task_data in TASK_MODELS.values():
        for model_id, model_name in task_data["options"].items():
            models[model_id] = model_name
    return models


def format_model_name(model_id):
    models = get_all_model_names()
    model_name = models.get(model_id, model_id)
    return f"{model_name} · {model_id}"


def call_llm(messages, model, temperature=0.3, max_tokens=1000):
    api_key = get_api_key()

    if not api_key:
        raise RuntimeError("دخل المفتاح الخاص بيك من Sidebar أو ضيفه في ملف .env باسم OPENROUTER_KEY")

    url = "https://openrouter.ai/api/v1/chat/completions"

    payload = {
        "model": model,
        "messages": messages,
        "temperature": temperature,
        "max_tokens": max_tokens,
    }

    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
        "HTTP-Referer": os.getenv("APP_URL", "http://localhost:8501"),
        "X-Title": APP_TITLE,
    }

    response = requests.post(url, json=payload, headers=headers, timeout=90)

    if response.status_code >= 400:
        try:
            error_data = response.json()
            error_message = error_data.get("error", {}).get("message", response.text)
        except Exception:
            error_message = response.text
        raise RuntimeError(f"OpenRouter Error {response.status_code}: {error_message}")

    data = response.json()

    try:
        return data["choices"][0]["message"]["content"]
    except Exception:
        raise RuntimeError(f"Unexpected API response: {data}")


def build_system_prompt(task):
    prompts = {
        "🤖 محادثة عامة": (
            "You are a friendly and helpful AI assistant. "
            "Answer clearly and naturally. "
            "Always reply in the same language the user is writing in."
        ),
        "💻 كود": (
            "You are an expert software engineer. "
            "Help with code, debugging, architecture, and technical questions. "
            "When you write code, keep code blocks clean and keep explanations separate. "
            "Always reply in the same language the user is writing in."
        ),
        "📝 تلخيص": (
            "You are a professional summarizer. "
            "Provide clear, structured summaries that capture the main points. "
            "Always reply in the same language the user is writing in."
        ),
        "🌐 ترجمة متخصصة": (
            "You are a professional translator with expertise in technical, legal, and medical terminology. "
            "Provide accurate and natural translations. "
            "Always reply in the same language the user is writing in."
        ),
        "🖼️ تحليل صورة": (
            "You are an expert image analyst. "
            "Describe and analyze images in detail. "
            "Always reply in the same language the user is writing in."
        ),
        "🎤 صوت": (
            "You are a helpful assistant that processes and responds to audio transcriptions. "
            "Always reply in the same language the user is writing in."
        ),
        "📊 تحليل بيانات": (
            "You are a data analysis expert. "
            "Analyze data, identify patterns, provide insights, and suggest visualizations. "
            "Always reply in the same language the user is writing in."
        ),
        "🎓 شرح وتعليم": (
            "You are an expert teacher who explains complex topics simply. "
            "Use examples, analogies, and step-by-step explanations. "
            "Always reply in the same language the user is writing in."
        ),
        "💼 CV وإيميلات": (
            "You are a professional career coach and business writer. "
            "Help write, improve, and format CVs and professional emails. "
            "Always reply in the same language the user is writing in."
        ),
        "🔍 بحث وتلخيص مقالات": (
            "You are a research assistant. "
            "Summarize articles, extract key insights, and provide structured research summaries. "
            "Always reply in the same language the user is writing in."
        ),
        "🧮 مسائل رياضية": (
            "You are a mathematics expert. "
            "Solve problems step by step, explain the solution clearly, and verify the answer. "
            "Always reply in the same language the user is writing in."
        ),
        "🏥 استشارة طبية": (
            "You are a medical information assistant. "
            "Provide general health information based on established medical knowledge. "
            "Always remind users to consult a real doctor for personal medical advice. "
            "Always reply in the same language the user is writing in."
        ),
        "⚖️ استشارة قانونية": (
            "You are a legal information assistant. "
            "Provide general legal information and explanations. "
            "Always remind users to consult a real lawyer for personal legal advice. "
            "Always reply in the same language the user is writing in."
        ),
    }

    return prompts.get(task, prompts["🤖 محادثة عامة"])


def validate_text_length(text):
    if len(text) > MAX_TEXT_CHARS:
        raise ValueError(f"الملف طويل جدًا. الحد الأقصى المسموح به هو {MAX_TEXT_CHARS} حرف.")


def process_file(uploaded_file):
    uploaded_file.seek(0)
    file_type = uploaded_file.type
    file_name = uploaded_file.name

    if file_type == "application/pdf":
        pdf_reader = PyPDF2.PdfReader(uploaded_file)
        page_count = len(pdf_reader.pages)

        if page_count > MAX_PDF_PAGES:
            raise ValueError(f"ملف PDF يحتوي على {page_count} صفحة. الحد الأقصى المسموح به هو {MAX_PDF_PAGES} صفحات.")

        text = ""
        for page in pdf_reader.pages:
            page_text = page.extract_text() or ""
            text += page_text + "\n"

        validate_text_length(text)
        return f"محتوى الملف ({file_name}):\n{text}"

    if file_type == "text/csv":
        df = pd.read_csv(uploaded_file)

        if len(df) > MAX_EXCEL_ROWS:
            raise ValueError(f"الملف يحتوي على {len(df)} صف. الحد الأقصى المسموح به هو {MAX_EXCEL_ROWS} صف.")

        text = df.to_string()
        validate_text_length(text)
        return f"محتوى الملف ({file_name}):\n{text}"

    if file_type in [
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "application/vnd.ms-excel",
    ]:
        sheets = pd.read_excel(uploaded_file, sheet_name=None)
        result = []

        for sheet_name, df in sheets.items():
            if len(df) > MAX_EXCEL_ROWS:
                raise ValueError(f"Sheet باسم {sheet_name} يحتوي على {len(df)} صف. الحد الأقصى المسموح به هو {MAX_EXCEL_ROWS} صف.")
            result.append(f"Sheet: {sheet_name}\n{df.to_string()}")

        text = "\n\n".join(result)
        validate_text_length(text)
        return f"محتوى الملف ({file_name}):\n{text}"

    if file_type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
        doc = Document(uploaded_file)
        paragraphs = [para.text for para in doc.paragraphs if para.text.strip()]

        if len(paragraphs) > MAX_DOCX_PARAGRAPHS:
            raise ValueError(f"ملف Word يحتوي على {len(paragraphs)} فقرة. الحد الأقصى المسموح به هو {MAX_DOCX_PARAGRAPHS} فقرة.")

        text = "\n".join(paragraphs)
        validate_text_length(text)
        return f"محتوى الملف ({file_name}):\n{text}"

    if file_type == "text/plain":
        text = uploaded_file.read().decode("utf-8", errors="ignore")
        validate_text_length(text)
        return f"محتوى الملف ({file_name}):\n{text}"

    raise ValueError(f"نوع الملف ({file_type}) مش مدعوم.")


def process_image(uploaded_image):
    uploaded_image.seek(0)
    image_data = uploaded_image.read()
    encoded_image = base64.b64encode(image_data).decode("utf-8")
    mime_type = uploaded_image.type
    return encoded_image, mime_type


def get_audio_format(mime_type):
    audio_format = mime_type.split("/")[-1] if "/" in mime_type else "wav"
    formats = {
        "x-wav": "wav",
        "wave": "wav",
        "mpeg": "mp3",
        "mpga": "mp3",
    }
    return formats.get(audio_format, audio_format)


def process_audio(uploaded_audio):
    uploaded_audio.seek(0)
    audio_data = uploaded_audio.read()
    encoded_audio = base64.b64encode(audio_data).decode("utf-8")
    mime_type = uploaded_audio.type
    audio_format = get_audio_format(mime_type)
    return encoded_audio, audio_format


def get_attachment_names(uploaded_file=None, uploaded_image=None, uploaded_audio=None):
    attachments = []

    if uploaded_file:
        attachments.append(f"📄 {uploaded_file.name}")

    if uploaded_image:
        attachments.append(f"🖼️ {uploaded_image.name}")

    if uploaded_audio:
        audio_name = getattr(uploaded_audio, "name", "recorded_audio")
        attachments.append(f"🎤 {audio_name}")

    return attachments


def build_user_content(user_input, model, uploaded_file=None, uploaded_image=None, uploaded_audio=None):
    text_parts = []

    if uploaded_file:
        text_parts.append(process_file(uploaded_file))

    text_parts.append(f"سؤال المستخدم:\n{user_input}")
    final_text = "\n\n".join(text_parts)

    if uploaded_audio and not model_supports_audio(model):
        raise ValueError(
            "الموديل الحالي لا يدعم إدخال الصوت. "
            "اختار Task الصوت أو موديل يدعم Audio Input، أو حوّل الصوت لنص وابعت النص."
        )

    if not uploaded_image and not uploaded_audio:
        return final_text

    user_content = []

    if uploaded_image:
        encoded_image, mime_type = process_image(uploaded_image)
        user_content.append(
            {
                "type": "image_url",
                "image_url": {
                    "url": f"data:{mime_type};base64,{encoded_image}"
                },
            }
        )

    if uploaded_audio:
        encoded_audio, audio_format = process_audio(uploaded_audio)
        user_content.append(
            {
                "type": "input_audio",
                "input_audio": {
                    "data": encoded_audio,
                    "format": audio_format,
                },
            }
        )

    user_content.append(
        {
            "type": "text",
            "text": final_text,
        }
    )

    return user_content


def build_messages(user_input, task, model, uploaded_file=None, uploaded_image=None, uploaded_audio=None, chat_history=None):
    messages = [
        {
            "role": "system",
            "content": build_system_prompt(task),
        }
    ]

    if chat_history:
        recent_history = chat_history[-MAX_HISTORY_MESSAGES:]
        for msg in recent_history:
            role = msg.get("role")
            content = msg.get("content")
            if role in ["user", "assistant"] and isinstance(content, str):
                messages.append(
                    {
                        "role": role,
                        "content": content,
                    }
                )

    messages.append(
        {
            "role": "user",
            "content": build_user_content(
                user_input=user_input,
                model=model,
                uploaded_file=uploaded_file,
                uploaded_image=uploaded_image,
                uploaded_audio=uploaded_audio,
            ),
        }
    )

    return messages


def generate_answer(
    user_input,
    task,
    model,
    temperature,
    max_tokens,
    uploaded_file=None,
    uploaded_image=None,
    uploaded_audio=None,
    chat_history=None,
):
    messages = build_messages(
        user_input=user_input,
        task=task,
        model=model,
        uploaded_file=uploaded_file,
        uploaded_image=uploaded_image,
        uploaded_audio=uploaded_audio,
        chat_history=chat_history,
    )

    return call_llm(
        messages=messages,
        model=model,
        temperature=temperature,
        max_tokens=max_tokens,
    )


def init_session_state():
    defaults = {
        "messages": [],
        "current_task": "🤖 محادثة عامة",
        "current_model": TASK_MODELS["🤖 محادثة عامة"]["default"],
        "quick_prompt": None,
        "manual_api_key": "",
    }

    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def reset_chat():
    st.session_state.messages = []
    st.session_state.quick_prompt = None


def main():
    st.set_page_config(
        page_title=APP_TITLE,
        page_icon="💬",
        layout="wide",
        initial_sidebar_state="expanded",
    )

    apply_custom_css()
    init_session_state()

    task, model, temperature, max_tokens, uploaded_file, uploaded_image, uploaded_audio = render_sidebar()

    render_header(task, model)
    render_chat_messages()

    user_input = st.chat_input("Type your question here...")

    if st.session_state.quick_prompt:
        user_input = st.session_state.quick_prompt
        st.session_state.quick_prompt = None

    if user_input:
        attachments = get_attachment_names(uploaded_file, uploaded_image, uploaded_audio)
        previous_messages = st.session_state.messages.copy()

        st.session_state.messages.append(
            {
                "role": "user",
                "content": user_input,
                "task": task,
                "model": model,
                "attachments": attachments,
            }
        )

        with st.chat_message("user", avatar="🧑‍💻"):
            st.markdown(user_input)

            if attachments:
                chips = "".join([f'<span class="attachment-chip">{item}</span>' for item in attachments])
                st.markdown(chips, unsafe_allow_html=True)

        try:
            with st.chat_message("assistant", avatar="🤖"):
                with st.spinner("Thinking..."):
                    answer = generate_answer(
                        user_input=user_input,
                        task=task,
                        model=model,
                        temperature=temperature,
                        max_tokens=max_tokens,
                        uploaded_file=uploaded_file,
                        uploaded_image=uploaded_image,
                        uploaded_audio=uploaded_audio,
                        chat_history=previous_messages,
                    )

                    st.markdown(answer)
                    st.markdown(
                        f"<div class='message-meta'>Task: {task} · Model: {model}</div>",
                        unsafe_allow_html=True,
                    )

            st.session_state.messages.append(
                {
                    "role": "assistant",
                    "content": answer,
                    "task": task,
                    "model": model,
                    "attachments": [],
                }
            )

        except Exception as e:
            error_message = f"Error: {str(e)}"
            st.error(error_message)

            st.session_state.messages.append(
                {
                    "role": "assistant",
                    "content": error_message,
                    "task": task,
                    "model": model,
                    "attachments": [],
                }
            )


def apply_custom_css():
    st.markdown(
        """
        <style>
        :root {
            --main-bg: #212121;
            --sidebar-bg: #171717;
            --card-bg: rgba(47, 47, 47, 0.78);
            --card-border: rgba(255, 255, 255, 0.10);
            --text-main: #ececec;
            --text-muted: #b4b4b4;
            --accent: #10a37f;
            --accent-soft: rgba(16, 163, 127, 0.14);
        }

        .stApp,
        [data-testid="stAppViewContainer"],
        [data-testid="stMain"] {
            background:
                radial-gradient(circle at top left, rgba(16, 163, 127, 0.10), transparent 28%),
                radial-gradient(circle at bottom right, rgba(52, 211, 153, 0.06), transparent 32%),
                var(--main-bg) !important;
            color: var(--text-main) !important;
        }

        header[data-testid="stHeader"] {
            background: rgba(33, 33, 33, 0.92) !important;
            border-bottom: 1px solid var(--card-border) !important;
        }

        section[data-testid="stSidebar"] {
            background: var(--sidebar-bg) !important;
            border-right: 1px solid var(--card-border) !important;
        }

        section[data-testid="stSidebar"] * {
            color: var(--text-main) !important;
        }

        .main .block-container,
        [data-testid="stMainBlockContainer"] {
            max-width: 1000px;
            padding-top: 1.4rem;
            padding-bottom: 7rem;
        }

        .app-header {
            background: linear-gradient(135deg, rgba(47,47,47,0.90), rgba(33,33,33,0.70)) !important;
            border: 1px solid var(--card-border) !important;
            border-radius: 26px !important;
            padding: 26px !important;
            margin-bottom: 18px !important;
            box-shadow: 0 18px 55px rgba(0,0,0,0.22) !important;
        }

        .app-badge {
            display: inline-flex;
            padding: 7px 12px;
            border-radius: 999px;
            background: var(--accent-soft) !important;
            border: 1px solid rgba(16, 163, 127, 0.25) !important;
            color: #a7f3d0 !important;
            font-size: 13px;
            font-weight: 700;
            margin-bottom: 14px;
        }

        .app-title {
            font-size: 38px;
            line-height: 1.1;
            font-weight: 800;
            letter-spacing: -1px;
            margin: 0;
            color: var(--text-main) !important;
        }

        .app-subtitle {
            color: var(--text-muted) !important;
            font-size: 15px;
            line-height: 1.7;
            margin-top: 10px;
        }

        .info-grid {
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 12px;
            margin-top: 18px;
        }

        .info-card {
            background: rgba(255,255,255,0.035) !important;
            border: 1px solid var(--card-border) !important;
            border-radius: 18px !important;
            padding: 13px 15px !important;
        }

        .info-label {
            color: var(--text-muted) !important;
            font-size: 12px;
            margin-bottom: 6px;
        }

        .info-value {
            color: var(--text-main) !important;
            font-size: 14px;
            font-weight: 700;
            word-break: break-word;
        }

        .sidebar-brand {
            padding: 17px 14px;
            border-radius: 22px;
            background: linear-gradient(135deg, rgba(16, 163, 127, 0.18), rgba(255,255,255,0.04)) !important;
            border: 1px solid var(--card-border) !important;
            margin-bottom: 16px;
        }

        .sidebar-title {
            font-size: 19px;
            font-weight: 800;
            color: var(--text-main) !important;
            margin-bottom: 6px;
        }

        .sidebar-caption {
            color: var(--text-muted) !important;
            font-size: 13px;
            line-height: 1.55;
        }

        .task-box {
            padding: 12px 14px;
            border-radius: 17px;
            background: rgba(16, 163, 127, 0.10) !important;
            border: 1px solid rgba(16, 163, 127, 0.20) !important;
            color: #d9fff2 !important;
            font-size: 13px;
            line-height: 1.65;
            margin-bottom: 10px;
        }

        .model-chip {
            display: inline-block;
            padding: 6px 9px;
            margin: 4px 4px 4px 0;
            border-radius: 999px;
            background: rgba(255,255,255,0.055) !important;
            border: 1px solid var(--card-border) !important;
            color: #eeeeee !important;
            font-size: 12px;
            line-height: 1.2;
        }

        .attachment-chip {
            display: inline-block;
            padding: 6px 10px;
            margin: 4px 4px 0 0;
            border-radius: 999px;
            background: rgba(16, 163, 127, 0.12) !important;
            border: 1px solid rgba(16, 163, 127, 0.25) !important;
            color: #bdf7e4 !important;
            font-size: 12px;
            font-weight: 600;
        }

        .empty-state {
            text-align: center;
            border: 1px dashed rgba(255,255,255,0.15) !important;
            background: rgba(255,255,255,0.035) !important;
            border-radius: 24px !important;
            padding: 34px 24px !important;
            margin-top: 16px;
        }

        .empty-title {
            color: var(--text-main) !important;
            font-size: 23px;
            font-weight: 800;
            margin-bottom: 8px;
        }

        .empty-subtitle {
            color: var(--text-muted) !important;
            font-size: 15px;
            line-height: 1.7;
        }

        div[data-testid="stChatMessage"] {
            border-radius: 22px !important;
            padding: 10px 14px !important;
            margin-bottom: 13px !important;
            background: rgba(255,255,255,0.035) !important;
            border: 1px solid rgba(255,255,255,0.075) !important;
        }

        div[data-testid="stChatMessage"] * {
            color: var(--text-main) !important;
        }

        div[data-testid="stChatMessage"] p,
        div[data-testid="stChatMessage"] li {
            line-height: 1.75;
        }

        .message-meta {
            color: var(--text-muted) !important;
            font-size: 12px;
            margin-top: 10px;
        }

        .stButton > button {
            border-radius: 15px !important;
            border: 1px solid var(--card-border) !important;
            background: rgba(255,255,255,0.055) !important;
            color: var(--text-main) !important;
            font-weight: 700 !important;
            transition: 0.18s ease;
        }

        .stButton > button:hover {
            border-color: rgba(16,163,127,0.45) !important;
            background: rgba(16,163,127,0.16) !important;
            color: #ffffff !important;
        }

        div[data-testid="stChatInput"] {
            background: rgba(33,33,33,0.76) !important;
            border-top: 1px solid var(--card-border) !important;
            backdrop-filter: blur(14px);
        }

        textarea,
        input,
        .stTextInput input,
        .stTextArea textarea {
            color: var(--text-main) !important;
            background: rgba(255,255,255,0.055) !important;
            border-color: var(--card-border) !important;
        }

        div[data-baseweb="select"] > div {
            background: rgba(255,255,255,0.055) !important;
            color: var(--text-main) !important;
            border-color: var(--card-border) !important;
        }

        .stFileUploader {
            background: rgba(255,255,255,0.035) !important;
            border-radius: 16px !important;
        }

        [data-testid="stAudioInput"] {
            background: rgba(255,255,255,0.055) !important;
            border: 1px solid var(--card-border) !important;
            border-radius: 16px !important;
            padding: 10px !important;
        }

        [data-testid="stAudioInput"] button {
            background: var(--accent) !important;
            color: white !important;
            border-radius: 999px !important;
            border: none !important;
        }

        .stAlert {
            border-radius: 16px !important;
        }

        @media (max-width: 768px) {
            .app-title {
                font-size: 29px;
            }

            .info-grid {
                grid-template-columns: 1fr;
            }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_header(task, model):
    model_name = TASK_MODELS[task]["options"].get(model, model)

    st.markdown(
        f"""
        <div class="app-header">
            <div class="app-badge">● OpenRouter AI Workspace</div>
            <h1 class="app-title">{APP_TITLE}</h1>
            <div class="app-subtitle">
                {APP_SUBTITLE}. اختر نوع المهمة والموديل المناسب من الـ Sidebar، وارفع ملفات أو صور أو صوت حسب احتياجك.
            </div>
            <div class="info-grid">
                <div class="info-card">
                    <div class="info-label">Current Task</div>
                    <div class="info-value">{task}</div>
                </div>
                <div class="info-card">
                    <div class="info-label">Active Model</div>
                    <div class="info-value">{model_name}</div>
                </div>
                <div class="info-card">
                    <div class="info-label">Memory</div>
                    <div class="info-value">Last {MAX_HISTORY_MESSAGES} messages</div>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_sidebar():
    with st.sidebar:
        st.markdown(
            f"""
            <div class="sidebar-brand">
                <div class="sidebar-title">{APP_TITLE}</div>
                <div class="sidebar-caption">Multi-task chatbot with clean task-based model selection.</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        st.markdown("### API Key")

        env_key_exists = bool(os.getenv("OPENROUTER_KEY") or os.getenv("OPENROUTER_API_KEY"))

        if env_key_exists:
            st.success("API Key loaded from .env")
        else:
            st.warning("دخل المفتاح الخاص بيك لتشغيل البوت.")

        st.session_state.manual_api_key = st.text_input(
            "OpenRouter API Key",
            value=st.session_state.get("manual_api_key", ""),
            type="password",
            placeholder="sk-or-v1-...",
            help="لو كتبت مفتاح هنا، هيستخدمه بدل المفتاح الموجود في .env",
        )

        st.markdown("---")
        st.header("Settings")

        tasks = list(TASK_MODELS.keys())

        task = st.selectbox(
            "Task Type",
            options=tasks,
            index=tasks.index(st.session_state.current_task),
        )

        if task != st.session_state.current_task:
            st.session_state.current_task = task
            st.session_state.current_model = TASK_MODELS[task]["default"]
            st.rerun()

        st.markdown(
            f"""
            <div class="task-box">
                <b>{task}</b><br>{TASK_DESCRIPTIONS.get(task, "")}
            </div>
            """,
            unsafe_allow_html=True,
        )

        model_options = list(TASK_MODELS[task]["options"].keys())

        if st.session_state.current_model not in model_options:
            st.session_state.current_model = TASK_MODELS[task]["default"]

        model = st.selectbox(
            "Available Models For This Task",
            options=model_options,
            index=model_options.index(st.session_state.current_model),
            format_func=format_model_name,
        )

        st.session_state.current_model = model

        temperature = st.slider(
            "Temperature",
            min_value=0.0,
            max_value=1.0,
            value=0.3,
            step=0.1,
        )

        max_tokens = st.slider(
            "Max Tokens",
            min_value=200,
            max_value=4000,
            value=1000,
            step=100,
        )

        st.markdown("---")

        uploaded_file = st.file_uploader(
            "Upload File",
            type=["pdf", "csv", "xlsx", "xls", "docx", "txt"],
        )

        st.caption(
            f"PDF بحد أقصى {MAX_PDF_PAGES} صفحات. Excel/CSV بحد أقصى {MAX_EXCEL_ROWS} صف. Word بحد أقصى {MAX_DOCX_PARAGRAPHS} فقرة. TXT بحد أقصى {MAX_TEXT_CHARS} حرف."
        )

        uploaded_image = st.file_uploader(
            "Upload Image",
            type=["png", "jpg", "jpeg", "webp"],
        )

        st.caption("ارفع صورة واضحة بصيغة PNG أو JPG أو JPEG أو WEBP.")

        uploaded_audio_file = st.file_uploader(
            "Upload Audio",
            type=["mp3", "wav", "m4a", "ogg", "webm"],
        )

        st.caption("ارفع ملف صوتي قصير وواضح. دعم الصوت يعتمد على الموديل المختار.")

        recorded_audio = None

        if hasattr(st, "audio_input"):
            recorded_audio = st.audio_input("Record Audio")
            st.caption("اضغط على زر الميكروفون داخل الصندوق واسمح للمتصفح باستخدام الميكروفون.")
        else:
            st.caption("تسجيل الصوت المباشر غير مدعوم في نسخة Streamlit الحالية. حدّث Streamlit.")

        uploaded_audio = recorded_audio or uploaded_audio_file

        if uploaded_audio and not model_supports_audio(model):
            st.warning("الموديل الحالي لا يدعم الصوت. اختار Task 🎤 صوت أو موديل يدعم Audio Input.")

        attachments = get_attachment_names(uploaded_file, uploaded_image, uploaded_audio)

        if attachments:
            chips = "".join([f'<span class="attachment-chip">{item}</span>' for item in attachments])
            st.markdown(chips, unsafe_allow_html=True)

        st.markdown("---")

        col1, col2 = st.columns(2)

        with col1:
            if st.button("New Chat", use_container_width=True):
                reset_chat()
                st.rerun()

        with col2:
            if st.button("Quick", use_container_width=True):
                st.session_state.quick_prompt = QUICK_PROMPTS.get(task, "ابدأ محادثة جديدة.")
                st.rerun()

        st.markdown("---")

        with st.expander("Used Models By Task", expanded=False):
            for task_name, task_data in TASK_MODELS.items():
                st.markdown(f"**{task_name}**")
                chips = "".join(
                    [
                        f'<span class="model-chip">{model_name}</span>'
                        for model_name in task_data["options"].values()
                    ]
                )
                st.markdown(chips, unsafe_allow_html=True)

        with st.expander("Current Model Details", expanded=False):
            st.markdown(f"**Task:** {task}")
            st.markdown(f"**Model Name:** {TASK_MODELS[task]['options'][model]}")
            st.code(model, language="text")

        st.caption(f"Developer: {APP_DEVELOPER}")

    return task, model, temperature, max_tokens, uploaded_file, uploaded_image, uploaded_audio


def render_empty_state():
    st.markdown(
        """
        <div class="empty-state">
            <div class="empty-title">ابدأ محادثتك الآن</div>
            <div class="empty-subtitle">
                اكتب رسالتك تحت، أو اختار Task من الـ Sidebar، أو ارفع ملف/صورة/صوت وخلي البوت يتعامل معاهم.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_chat_messages():
    if not st.session_state.messages:
        render_empty_state()
        return

    for msg in st.session_state.messages:
        role = msg.get("role", "assistant")
        content = msg.get("content", "")
        task = msg.get("task")
        model = msg.get("model")
        attachments = msg.get("attachments", [])

        avatar = "🧑‍💻" if role == "user" else "🤖"

        with st.chat_message(role, avatar=avatar):
            st.markdown(content)

            if attachments:
                chips = "".join([f'<span class="attachment-chip">{item}</span>' for item in attachments])
                st.markdown(chips, unsafe_allow_html=True)

            if role == "assistant" and task and model:
                st.markdown(
                    f"<div class='message-meta'>Task: {task} · Model: {model}</div>",
                    unsafe_allow_html=True,
                )


if __name__ == "__main__":
    main()