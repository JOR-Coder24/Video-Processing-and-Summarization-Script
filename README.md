# Video Processing and Summarization Script

This Python script processes video files to extract audio, convert speech to text, generate summaries of the transcriptions, and save the results in a Word document with annotated frames. It supports processing both individual video files and entire folders containing videos. The following steps outline how the script works:

---

## Features

- **Extract Audio from Video**: The script extracts audio from video files (supports formats such as `.mp4`, `.avi`, `.mov`, `.mkv`).
- **Speech to Text**: The audio is transcribed to text in 60-second segments using Google's Speech-to-Text API.
- **Speech Annotation with Timestamps**: Each transcribed segment is time-stamped to match the exact moment in the video.
- **Summarization**: The transcribed text is summarized using the BART model from the Hugging Face library.
- **Frame Extraction**: Key frames are extracted from the video and attached below the corresponding summaries in a Word document.
- **Word Document Output**: A Word document is created containing the summarized text and associated video frames.

---

## Requirements

To use this script, you'll need the following Python packages:

- **moviepy**: For video file processing.
- **speech_recognition**: For speech-to-text conversion.
- **transformers**: For using Hugging Face's pre-trained BART model for text summarization.
- **docx**: For generating Word documents.
- **PIL (Pillow)**: For image handling and saving video frames.

Install these dependencies with the following command:

```bash
pip install moviepy SpeechRecognition transformers python-docx pillow
```

---

## Usage

### Running the Script

To run the script, simply execute the Python file from the command line:

```bash
python video_processing.py
```

The script will prompt you to choose one of the following options:

1. **Process a single video file**: Allows you to process one video file at a time.
2. **Process multiple video files in a folder**: Processes all video files within a folder.
3. **Quit**: Exits the script.

### Processing a Single Video File

If you choose the **file** option, you will be prompted to enter the full path of the video file you want to process. For example:

```
Enter the video file path: /path/to/video.mp4
```

The script will extract the audio from the video, transcribe it to text, summarize it, and save the results in a Word document with corresponding frames.

### Processing Multiple Video Files in a Folder

If you choose the **folder** option, you will be asked to provide the folder path that contains the video files. For example:

```
Enter the folder path with video files: /path/to/folder
```

The script will process all video files in the folder that match common video file extensions (`.mp4`, `.avi`, `.mov`, `.mkv`), extracting audio, transcribing, summarizing, and saving results for each video.

### Exiting the Script

To exit the program, simply enter **quit** when prompted.

---

## Output

For each processed video, the script generates a Word document with the following contents:

1. **Summary of the video**: Each paragraph from the transcribed text is summarized.
2. **Timestamps**: Each summary includes the timestamp of the segment it corresponds to.
3. **Video frames**: A key frame from the video is included beneath each summary, representing the corresponding moment in the video.

The Word document will be named using the base filename of the video followed by `_summarized_with_images.docx`.

For example, for a video named `example_video.mp4`, the output will be `example_video_summarized_with_images.docx`.

---

## Example Output

```plaintext
Summary of video: example_video

Summary of Paragraph 1 [00:00:00]:
The video introduces the topic of machine learning and its applications in various industries.

[Image Frame 1]

Summary of Paragraph 2 [00:01:00]:
The speaker goes on to explain the different types of machine learning algorithms and their differences.

[Image Frame 2]
```

---

## Cleanup

- Temporary audio files are created during processing and automatically deleted after use.
- Temporary image files are also deleted after they are added to the Word document.

---

## Notes

- Ensure that the video files are in supported formats (e.g., `.mp4`, `.avi`, `.mov`, `.mkv`).
- The video must have audio in a recognizable language for the speech-to-text process to work.
- The BART summarization model may require a stable internet connection for the first run if the model has not been cached.
- You may need to install additional dependencies or libraries depending on your environment.

---

## Troubleshooting

- **Error in speech recognition**: Ensure that the video has clear and audible speech. If the speech is unclear or distorted, the recognition may fail.
- **Word document not saving**: Ensure that the Python script has write permissions to the folder where you're saving the output document.

---

## License

This script is free to use and modify for personal or educational purposes. If you decide to redistribute it, please include this README.

