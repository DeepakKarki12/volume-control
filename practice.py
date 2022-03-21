import cv2 as cv
import mediapipe as mp
import math
import numpy
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(
IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))
volume.GetVolumeRange()  # the range of volume is (-65,0)

mphands = mp.solutions.hands
hands = mphands.Hands()
draw = mp.solutions.drawing_utils

cap = cv.VideoCapture(0)

while 1:

    _, img = cap.read()
    h, w, c = img.shape
    img_rgb = cv.cvtColor(img, cv.COLOR_BGR2RGB)  # mediapipe cannot process the bgr image so converted it -> rgb
    result = hands.process(img_rgb)  # mp processing the rgb img
    if result.multi_hand_landmarks:  # if there are multiple hands in result
        for single_hand in result.multi_hand_landmarks:  # derived single hand from multiple hand
            draw.draw_landmarks(img, single_hand,
                                mphands.HAND_CONNECTIONS)  # drawing the landmarks and interconnecting them
            for id, lm in enumerate(
                    single_hand.landmark):  # enumerate gives a tuple which combine the output with there id's
                x, y = int(lm.x * w), int(lm.y * h)  # we need to convert the x and y co-ordinate in pixel no.
                if id == 4:  # id of tip of thumb
                    in_x, in_y = x, y
                if id == 8:  # id of tip of index
                    th_x, th_y = x, y
            cv.line(img, (in_x, in_y), (th_x, th_y), (0, 0, 0), 3, cv.FILLED)  # drawing line b/w index tip and thumb tip
            diss = math.hypot(th_x-in_x,th_y-in_y)  # distance b/w index tip and thumb tip
            diss = numpy.interp(diss, (40, 180), (-65, 0))  # converting the range of diss into desired range
            print(diss)  # printing the distance
            volume.SetMasterVolumeLevel(diss, None)  # changing the volume of computer

    cv.imshow("window", img)
    if cv.waitKey(1) == ord('q'):
        break
cap.release()
cv.destroyAllWindows()
