import win32com.client
from cvzone.HandTrackingModule import HandDetector
import cv2
import os

# Get the absolute path to the file
file_path = os.path.join(os.path.dirname(__file__), "Sushil-Presentation.pptx")
print(f"Attempting to open file: {file_path}")

# Initialize PowerPoint Application
try:
    Application = win32com.client.Dispatch("PowerPoint.Application")
    Presentation = Application.Presentations.Open(file_path)
    print(f"Presentation '{Presentation.Name}' opened successfully.")
    Presentation.SlideShowSettings.Run()
except Exception as e:
    print(f"Error opening PowerPoint file: {e}")
    exit()

# Parameters
width, height = 900, 720
gestureThreshold = 300

# Camera Setup
cap = cv2.VideoCapture(0)
cap.set(3, width)
cap.set(4, height)

# Hand Detector
detectorHand = HandDetector(detectionCon=0.8, maxHands=1)

# Variables
imgList = []
delay = 30
buttonPressed = False
counter = 0
drawMode = False
imgNumber = 20
delayCounter = 0
annotations = [[]]
annotationNumber = -1
annotationStart = False

while True:
    # Get image frame
    success, img = cap.read()
    if not success:
        print("Failed to capture image")
        break

    # Find the hand and its landmarks
    hands, img = detectorHand.findHands(img)  # with draw
    if hands and buttonPressed is False:  # If hand is detected
        hand = hands[0]
        cx, cy = hand["center"]
        lmList = hand["lmList"]  # List of 21 Landmark points
        fingers = detectorHand.fingersUp(hand)  # List of which fingers are up
        if cy <= gestureThreshold:  # If hand is at the height of the face
            if fingers == [1, 1, 1, 1, 1]:
                print("Next")
                buttonPressed = True
                if imgNumber > 0:
                    Presentation.SlideShowWindow.View.Next()
                    imgNumber -= 1
                    annotations = [[]]
                    annotationNumber = -1
                    annotationStart = False
            if fingers == [1, 0, 0, 0, 0]:
                print("Previous")
                buttonPressed = True
                if imgNumber > 0:
                    Presentation.SlideShowWindow.View.Previous()
                    imgNumber += 1
                    annotations = [[]]
                    annotationNumber = -1
                    annotationStart = False

    else:
        annotationStart = False

    if buttonPressed:
        counter += 1
        if counter > delay:
            counter = 0
            buttonPressed = False

    for i, annotation in enumerate(annotations):
        for j in range(len(annotation)):
            if j != 0:
                cv2.line(img, annotation[j - 1], annotation[j], (0, 0, 200), 12)

    cv2.imshow("Image", img)

    key = cv2.waitKey(1)
    if key == ord('q'):
        break

# Release resources
cap.release()
cv2.destroyAllWindows()
Presentation.Close()
Application.Quit()