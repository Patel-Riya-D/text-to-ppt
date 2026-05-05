# Text to PPT

Text to PPT is an AI-powered presentation builder that turns a topic into a slide outline and a downloadable PowerPoint deck.

The project has two parts:

- `app.py`: a Streamlit chat-style UI for creating and editing presentations
- `backend/main.py`: a FastAPI service that generates outlines and builds PPT files

## What it does

- Generates slide outlines from a topic
- Builds PowerPoint decks with clean, presentation-ready content
- Supports tone-based output such as `Professional`
- Accepts optional logo and content image uploads
- Sanitizes LLM output so markdown, icon tokens, and emoji do not leak into slides
- Applies theme colors and presentation layouts during deck creation

## Tech Stack

- Python
- Streamlit
- FastAPI
- Azure OpenAI
- `python-pptx`

## Project Structure

```text
.
├── app.py
├── backend
│   ├── config.py
│   ├── main.py
│   └── services
│       └── ppt_service.py
        └── intelligence_layer.py
        └── intent_classifier.py
├── requirements.txt
└── README.md
```

## Requirements

- Python 3.10 or newer
- Access to an Azure OpenAI deployment

## Setup

1. Create and activate a virtual environment.

```bash
python -m venv venv
source venv/bin/activate
```

On Windows:

```bash
venv\Scripts\activate
```

2. Install dependencies.

```bash
pip install -r requirements.txt
```

3. Add the required environment variables.


## Run the App

The project expects the FastAPI backend to run on port `9000`.

1. Start the backend:

```bash
uvicorn backend.main:app --reload --port 9000
```

2. In a second terminal, start the Streamlit UI:

```bash
streamlit run app.py
```

3. Open the Streamlit app in your browser and generate a presentation.

## API Endpoints

### `GET /`

Health check for the backend.

Response:

```json
{"message":"FastAPI backend running 🚀"}
```

### `POST /generate-outline`

Generates slide content for a topic.

Form fields:

- `topic` - required
- `num_slides` - optional, default `5`
- `tone` - optional, default `Professional`
- `theme_colors` - optional JSON string

### `POST /build-ppt`

Builds the PowerPoint file from generated slide data.

Form fields:

- `topic` - required
- `tone` - optional, default `Professional`
- `slides_json` - required JSON string containing slide data
- `logo` - optional file upload

The response includes:

- the cleaned slide list
- a base64-encoded PPT file

## How It Works

1. The UI sends a topic to the backend
2. The backend uses Azure OpenAI to generate an outline
3. Slide content is cleaned and normalized
4. The service renders the final PowerPoint deck with `python-pptx`

The generation pipeline is intentionally defensive:

- flatten AI output
- clean text
- validate content
- render slides

This keeps icons, markdown, and noisy tokens out of the final presentation.

## Notes

- Uploaded files/logo are stored in an `uploads/` folder during processing

