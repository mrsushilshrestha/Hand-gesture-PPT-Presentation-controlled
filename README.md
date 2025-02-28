# Hand Gesture Controlled PowerPoint Presentation

## Overview
This project enables users to control a PowerPoint presentation using hand gestures detected via a webcam. The system utilizes OpenCV for camera input, cvzone's HandTrackingModule for gesture recognition, and the `pywin32` library to interact with Microsoft PowerPoint.

## Features
- 🎥 **Real-time Hand Gesture Recognition**: Uses OpenCV and cvzone for accurate hand tracking.
- 📽️ **PowerPoint Slideshow Control**: Navigate through slides using simple hand gestures.
- 🤚 **Gesture Commands**:
  - ✋ **All fingers up** → Move to the next slide.
  - 👎 **Only index finger down** → Move to the previous slide.
- 🔄 **Automatic Slideshow Execution**: Opens and starts the PowerPoint presentation automatically.
- 🖥️ **User-Friendly Interface**: Simple and intuitive controls.

## Installation
Ensure you have Python installed along with the following dependencies:

```bash
pip install opencv-python cvzone pywin32
```

## Usage
1. **Prepare the PowerPoint File**:
   - Place `Sushil-Presentation.pptx` in the same directory as the script.

2. **Run the Script**:
   ```bash
   python script.py
   ```

3. **Control the Presentation with Hand Gestures**:
   - ✋ **All fingers up** → Next slide
   - ☝️ **Only index finger up** → Previous slide

4. **Exit the Program**:
   - Press 'q' to terminate the application.

## Prerequisites
- Microsoft PowerPoint must be installed on your system.
- A functional webcam is required for gesture detection.
- Adjust detection confidence or gesture thresholds if needed for better performance.

## Notes
- The script uses OpenCV for image capture and processing.
- The `pywin32` library is essential for interacting with PowerPoint.
- Works best in well-lit environments for accurate hand detection.

## License
This project is open-source and free to use for educational purposes.

## Contributors
- **Sushil Shrestha** - Developer

For contributions or improvements, feel free to fork this repository and submit a pull request. 🚀

