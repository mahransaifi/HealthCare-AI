import cv2
# import numpy as np
import win32com.client
import torch
import torchvision.transforms as transforms
from torchvision import models
from PIL import Image
import json
import requests

speaker = win32com.client.Dispatch("SAPI.SpVoice")


model = models.mobilenet_v2(weights=models.MobileNet_V2_Weights.DEFAULT)
model.eval()


LABELS_URL = "https://storage.googleapis.com/download.tensorflow.org/data/imagenet_class_index.json"
labels = json.loads(requests.get(LABELS_URL).text)


preprocess = transforms.Compose([
    transforms.Resize(256),
    transforms.CenterCrop(224),
    transforms.ToTensor(),
    transforms.Normalize(mean=[0.485, 0.456, 0.406], std=[0.229, 0.224, 0.225]),
])


cap = cv2.VideoCapture(0)

while True:
    ret, frame = cap.read()
    if not ret:
        break


    pil_img = Image.fromarray(cv2.cvtColor(frame, cv2.COLOR_BGR2RGB))
    img_tensor = preprocess(pil_img)
    img_tensor = img_tensor.unsqueeze(0)


    with torch.no_grad():
        preds = model(img_tensor)


    probs = torch.nn.functional.softmax(preds[0], dim=0)


    top_probs, top_indices = torch.topk(probs, 3)

    for i in range(3):
        index = top_indices[i].item()
        score = top_probs[i].item()
        label = labels[str(index)][1]

        print(f"{label}: {score:.2f}")

        if "watch" in label.lower() and score > 0.2:
            print("⚠️ Warning: digital_clock detected!")
            speaker.Speak("This is a digital watch of casio with 5 alarms and water resistant. It also shows world time")

        elif "bottle" in label.lower() and score > 0.2:
            print("⚠️ Warning: bottle detected detected!")
            speaker.Speak("This is a water bottle made with plastic")

        elif "banana" in label.lower() and score > 0.2:
            print("⚠️ Warning: banana detected!")
            speaker.Speak("This is banana. Rich in nutrients, Boost Energy, Aids Digestion, Heart Health, Good for weight Management and also improves mood")

        elif "peanut" in label.lower() and score > 0.2:
            print("⚠️ Warning: peanut detected!")
            speaker.Speak("This is a peanut, if you are suffering from peanut allergy it will be harmful for you as Severe Allergic Reactions,Skin Reactions,Digestive Issues, Respiratory Problems,Swelling")

        elif "apple" in label.lower() and score > 0.2:
            print("⚠️ Warning: Apple detected!")
            speaker.Speak("This is Apple. It is good for your health as it is Rich in Nutrients, good for Digestion, Heart Health, Aids in Weight Loss, Boosts Immunity")

    cv2.imshow("AI Assistant - Detector", frame)


    if cv2.waitKey(1) & 0xFF == ord('q'):
        break

cap.release()
cv2.destroyAllWindows()
