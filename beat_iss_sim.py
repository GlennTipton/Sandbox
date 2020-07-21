#disclaimer
#recently, it looks like i am not being able to dock the spaceship even though my rate is well under -0.2m/s 
#dunno if there is any SW flag that limits the fun of those who play it too often
#or if they wanted to scare developers off
#i decided to do my "release" version assuming the hard limit of -0.2, rather than anything the game decides to show as blue!

#how to run it: open the game, start this program, switch to the game
#see how it fares and press ESC when you're done
#assuming MS Edge window @1920x1080 and 250% zoom
#
#feel free to contact me for suggestions
#i am not an OO guy, i use lots of globals, i admit it is ugly
#feel free to teach me how to make it elegant

import pytesseract
import cv2
import numpy as np
from PIL import Image, ImageGrab, ImageEnhance
import win32api
import win32con
import threading
import time
import sched
import re
import os

#there is really *a lot* of interesting things to do with this game
#there goes a nonexhaustive list of things i may try someday
# - retrain tesseract with the characters seen in the game
# - fancy input filtering (which is rarely needed)

#assuming you are using windows
#you must download tesseract
#i strongly recommend tesseract 3, i had trouble with 4 and 5
pytesseract.pytesseract.tesseract_cmd = os.path.join(os.environ["localappdata"], 
                                                     "Tesseract-OCR", "tesseract.exe")

#all masks that we need to get green, blue, orange and red numbers
cool_red = [np.array([115,100,100]), np.array([135,255,255])]
red_orange_blue = [np.array([0,100,100]), np.array([25,255,255])]
green = [np.array([60,100,100]), np.array([70,255,255])]

#i know it is just so ugly to hard-code number positions
#but i want to scan just part of the window to save time
#i used 250% zoom in order to have nicer numbers
#according to my tests, things don't speed up with partial screenshots
#so i scan the entire number area
#(x top left, y top left, x bottom right, y bottom right)
all_numbers = (500, 150, 1410, 1000)

#i am not using angle rates
#computing them is safer, as we avoid transport delay
#reading them would be useful if we wanted to start the game randomly
#(i.e. if we started with nonzero angle rates)
#input order: x, y, z, roll, yaw, pitch, rate
#coords: (left, top, width, height)
inputs_positions = [
    (0, 350, 160, 40),
    (0, 390, 160, 40),
    (0, 430, 160, 40),
    (360, 0, 190, 50),
    (360, 730, 190, 50),
    (730, 360, 190, 50),
    (600, 660, 150, 40)
]

#input order: x, y, z, roll, yaw, pitch, rate
current_values = [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]
#i don't use the "rate" past value or integration error
past_values = [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]
integrated_errors = [0.0, 0.0, 0.0, 0.0, 0.0, 0.0]
rates = [0.0, 0.0, 0.0, 0.0, 0.0, 0.0]
ctrl_delays = [25, 25, 25, 2, 2, 2]

#control law for x axis is more complicated
#i shall have no overshoot and a very, very gentle approach near 0
#acceleration may eat cpu due to extensive i/o + long distance
#the game's rate evaluation is something i cannot understand at all
#if we approach fast, it goes red, no matter how slow we get
#grr! this drives me crazy! they didn't put a requirement stating this
#anyway, each input has four arrays:
#- breakpoints for interpolation
#- proportional gains
#- integral gains
# derivative gains
#gains are calculated by means of linear interpolation
gain_schedules = [
    [
        [-200.0, -90.0, -7.5, -3.0, 0.0, 3.0, 7.5, 90.0, 200.0],
        [0.2, 0.4, 0.7, 0.7, 0.0, 0.7, 0.7, 0.4, 0.2],
        [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0],
        [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]
     ],
    [
        [-6.0, -3.0, 0.0, 3.0, 6.0],
        [1.2, 1.2, 1.2, 1.2, 1.2],
        [0.0, 0.0, 0.02, 0.0, 0.0],
        [0.0, 0.0, 0.002, 0.0, 0.0]
    ],
    [
        [-6.0, -3.0, 0.0, 3.0, 6.0],
        [1.2, 1.2, 1.2, 1.2, 1.2],
        [0.0, 0.0, 0.02, 0.0, 0.0],
        [0.0, 0.0, 0.002, 0.0, 0.0]
    ],
    [
        [-180.0, -1.0, -0.5, 0.5, 1.0, 180.0],
        [0.2, 0.2, 0.2, 0.2, 0.2, 0.2],
        [0.0, 0.0, 0.01, 0.01, 0.0, 0.0],
        [0.0, 0.0, 0.0, 0.0, 0.0, 0.0]
    ], 
    [
        [-180.0, -1.0, -0.5, 0.5, 1.0, 180.0],
        [0.2, 0.2, 0.2, 0.2, 0.2, 0.2],
        [0.0, 0.0, 0.01, 0.01, 0.0, 0.0],
        [0.0, 0.0, 0.0, 0.0, 0.0, 0.0]
    ],
    [
        [-180.0, -1.0, -0.5, 0.5, 1.0, 180.0],
        [0.2, 0.2, 0.2, 0.2, 0.2, 0.2],
        [0.0, 0.0, 0.01, 0.01, 0.0, 0.0],
        [0.0, 0.0, 0.0, 0.0, 0.0, 0.0]
    ]
 ]

#mapping all keyboard codes we need
#input order: x, y, z, roll, yaw, pitch
increase_buttons = [0x51, 0x44, 0x57, 0x67, 0x64, 0x68]
decrease_buttons = [0x45, 0x41, 0x53, 0x69, 0x66, 0x65]

#flag used to kill the threads when we press "ESC"
keep_alive = True

#this function may be useful for you, but i do not use it
#mouse-clicking is slower than using the keyboard
#and i prefer to have control over it in case the program goes nuts
def click(coords):
    win32api.SetCursorPos(coords)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,coords[0],coords[1],0,0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,coords[0],coords[1],0,0)

def box_inside_box(box1, box2):
    #see if the object in area "box1" fits inside object in area "box2"
    return (
            (box1[0] >= box2[0])
            and (box1[1] >= box2[1])
            and (box1[0] + box1[2] <= box2[0] + box2[2])
            and (box1[1] + box1[3] <= box2[1] + box2[3])
    )

def consolidate_inputs(screenshot):
    #tesseract data (when ok) has 12 columns
    #the important stuff is the four coordinates and the value itself
    #coordinates are given as (left, top, width, height)
    screenshot = screenshot.split("\n")[1:]
    filtered_data = [] #only meaningful data from tesseract
    #in rare cases a number may be split. we try to concatenate its parts
    #each array element will contain all parts that belong to an input
    #e.g. "-" and "10.0" inside the area of the same input
    filtered_lines_per_input = [[], [], [], [], [], [], []]
    for line in screenshot:
        columns = line.split("\t")
        if len(columns) < 12:
            #this means tesseract could not read correctly
            #only angle rates can be estimated precisely
            #but i usually do not need these estimations
            #so, if a sensor is unreliable, i keep previous value
            return
        if (columns[11] == "text"):
            #first line (only labels)
            continue
        if (columns[10] == "-1"):
            #line without meaningful data (confidence = -1)
            continue
        #all remaining data is part of a possibly valid line
        #columns 6-9 are the coordinates
        #column 11 is the value
        filtered_data.append(columns[6:])
    #in the next for, i append all info for each desired input
    for i in range(0, len(filtered_lines_per_input)):
        for j in range(0, len(filtered_data)):
            #first, check if data belongs to one of the inputs i want
            #it may be one of the angle rates, or even garbage
            box = (int(filtered_data[j][0]),
                   int(filtered_data[j][1]),
                   int(filtered_data[j][2]),
                   int(filtered_data[j][3]))
            #looking for concatenated data (all fit in the same box)
            if box_inside_box(box, inputs_positions[i]):
                filtered_lines_per_input[i].append(j)
        #now, we can look at filtered_lines_per_input
        #and try to assemble the actual input
        new_input = ""
        #finally, concatenating (usually only one line)
        for j in range(0, len(filtered_lines_per_input[i])):
            new_input += filtered_data[filtered_lines_per_input[i][j]][5]
        #removing all unit symbols and whitespace
        new_input = new_input.replace("°", "").replace("m", "")
        new_input = new_input.replace("/", "").replace("s", "")
        new_input = new_input.replace(" ", "")
        #good case: a healthy number
        if re.fullmatch("-{0,1}[0-9]+\.[0-9]", new_input):
            new_input = float(new_input)        
        #a missing dot can be replaced
        #in rate, we have 3 decimal digits
        #in the others, we have one
        elif re.fullmatch("-{0,1}[0-9]+[0-9]", new_input):
            if i == 6:
                new_input = float(new_input)*0.001
            else:
                new_input = float(new_input)*0.1
        #bad case: we cannot recognize a number
        else:
            new_input = current_values[i]
        current_values[i] = new_input
        
def capture_value(box, masks):
    image = ImageGrab.grab(bbox=box)
    #image need to be in HSV to enable threshold over colors
    image = cv2.cvtColor(np.asarray(image), cv2.COLOR_BGR2HSV)
    #ORing all color masks to get any known number color
    mask = cv2.inRange(image, masks[0][0], masks[0][1])
    for i in range(0, len(masks)):
        mask = mask + cv2.inRange(image, masks[i][0], masks[i][1])
    image = cv2.bitwise_and(image, image, mask=mask)
    #converting image to BW so that tesseract works better (i hope)
    image = cv2.cvtColor(image, cv2.COLOR_HSV2BGR)
    image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    ret,image = cv2.threshold(image,10,255,cv2.THRESH_BINARY)
    image = Image.fromarray(image)
    #finally, the heavy job
    data = pytesseract.image_to_data(image, config="iss_sim")
    return data
    
#the control law will be a PID (and i do not even use the D often)
#for x axis, i check rate to ensure it will dock smoothly
def control_law(basic_input, past_input, rate_input, old_integrated_error,
                increase_btn, reduce_btn, Kp, Ki, Kd, special_x_axis=False):
    #a sort of anti-windup
    if Ki != 0.0:
        integrated_error = old_integrated_error - basic_input
    else:
        integrated_error = 0.0        
    basic_error = -basic_input
    #using some fix x-axis rate when close enough
    #this variable is updated very slowly in the game (once in 3 seconds?)
    #controlling it would be depressing
    if (special_x_axis and abs(basic_error) < 2.0):
        control_effort = -1.0
    elif (special_x_axis and abs(basic_error) < 5.0):
        control_effort = -2.0
    #otherwise, a kind of regular PID
    #tuned by trial and error (shame on me)
    else:
        control_effort = (basic_error*Kp
                          + integrated_error*Ki
                          + (basic_input - past_input)*Kd)
    #angle steps are always 0.1
    #since the game does not state a clear step for the axes,
    #i use arbitrarily 0.1 (it is actually lower) to keep commonality
    #then i use the gains to make it work properly
    if (control_effort > rate_input):
        button = increase_btn
        delta = 0.1
    else:
        button = reduce_btn
        delta = -0.1
    #control error = distance between desired and actual control action
    control_error = abs(round(10.0*(control_effort - rate_input)))
    #will limit controller action to keep threads within timing bounds
    for i in range (0, min(10, control_error)):
        #it looks like keybd_event is deprecated, but it still works
        #sendInput() is more modern but more laborious
        win32api.keybd_event(button, 0,0,0)
        #dunno if this time.sleep is strictly necessary
        time.sleep(0.01)
        win32api.keybd_event(button, 0, win32con.KEYEVENTF_KEYUP, 0)
        rate_input += delta
    return integrated_error, rate_input
    
def gain_scheduler(basic_input, gain_set):
    #good ol' linear interpolation to calculate the gains
    #did not perform binary array search, though i could
    #won't optimize microseconds when OCR takes 1 second =)
    last_idx = len(gain_set[0])-1
    if basic_input <= gain_set[0][0]:
        return gain_set[1][0], gain_set[2][0], gain_set[3][0]
    elif basic_input >= gain_set[0][last_idx]:
        return gain_set[1][last_idx], gain_set[2][last_idx], gain_set[3][last_idx]
    else:
        idx = 0
        while(basic_input >= gain_set[0][idx+1]):
            idx += 1
        interpolation_ratio = ((basic_input - gain_set[0][idx])
                               / (gain_set[0][idx+1] - gain_set[0][idx]))
        Kp = (gain_set[1][idx] 
              + (gain_set[1][idx+1] - gain_set[1][idx]) * interpolation_ratio)
        Ki = (gain_set[2][idx]
              + (gain_set[2][idx+1] - gain_set[2][idx]) * interpolation_ratio)
        Kd = (gain_set[3][idx]
              + (gain_set[3][idx+1] - gain_set[3][idx]) * interpolation_ratio)
        return Kp, Ki, Kd

class readThread (threading.Thread):
    def __init__(self, log, flag):
        threading.Thread.__init__(self)
        self.log = log
        self.flag = flag
        
    def run(self):
        #will write into memory only when app closes
        #assuming you won't leave it open till you run out of ram
        #check if you are overruning frame time
        #this will certainly degrade controller quality
        task_times = []
        while(keep_alive):
            #trying to make it behave like real-time
            self.flag.wait()
            task_start = time.time()
            self.flag.clear()
            #do the reading job
            screenshot = capture_value(all_numbers, [cool_red, red_orange_blue, green])
            #maybe i should write in the global only at thread end
            #or make each thread write in different globals
            #there is a risk we'll read out-of-order
            consolidate_inputs(screenshot)
            task_times.append(time.time()-task_start)
        for task_time in task_times:
            self.log.write("Frame CPU load: {:.1%}".format(task_time/2.4) + "\n")
        self.log.close()

class ctrlThread(threading.Thread):
    def __init__(self, log, index, flag, trigger):
        threading.Thread.__init__(self)
        self.flag = flag
        self.counter = 0
        self.index = index
        self.trigger = trigger
        self.log = log
        
    def run(self):
        #will write into memory only when app closes
        #assuming you won't leave it open till you run out of ram
        task_times = []
        while(keep_alive):
            #trying to make it behave like real-time
            self.flag.wait()
            task_start = time.time()
            self.flag.clear()
            Kp = 0.0
            Ki = 0.0
            Kd = 0.0
            #will start the controller only after a trigger
            #this is a lazy approach, but works
            #first i put the craft more or less in the right heading
            #then i can start controlling x, y and z
            #those really good with control systems may do fancier controllers
            if self.counter >= self.trigger:
                Kp, Ki, Kd = gain_scheduler(current_values[self.index],
                                            gain_schedules[self.index])
                #added rate control for x axis
                if (self.index == 0):
                    special_x_axis = True
                else:
                    special_x_axis = False
                integrated_errors[self.index], rates[self.index] = control_law(
                    current_values[self.index],
                    past_values[self.index],
                    rates[self.index],
                    integrated_errors[self.index], 
                    increase_buttons[self.index],
                    decrease_buttons[self.index], Kp, Ki, Kd, special_x_axis)
            for i in range(0, len(past_values)):
                past_values[self.index] = current_values[self.index]
            self.counter += 1
            task_times.append(time.time()-task_start)
        #control threads run relatively fast
        #but i suspect we do not notice when I/O is overloaded
        for task_time in task_times:
            self.log.write("Frame CPU load: {:.1%}".format(task_time/0.6)
                           + "\n")
        self.log.close()

if __name__ == '__main__':
    input("Hello! Press <ENTER> to start.")
    print("Switch to the game window! Quick!")
    time.sleep(3)
    #trying to imitate real-time scheduling
    s = sched.scheduler(time.time, time.sleep)
    ev_read_inputs = []
    read_logs = []
    read_threads = []
    for i in range(0, 4):
        ev_read_inputs.append(threading.Event())
        read_logs.append(open("log_read{0}.txt".format(str(i)), "w"))
        read_threads.append(readThread(read_logs[i], ev_read_inputs[i]))
        read_threads[i].start()   
    ev_ctrl = []
    ctrl_logs = []
    ctrl_threads = [] 
    for i in range(0, 6):
        ev_ctrl.append(threading.Event())
        ctrl_logs.append(open("log_ctrl{0}.txt".format(str(i)), "w"))
        ctrl_threads.append(ctrlThread(ctrl_logs[i],
                                       i, ev_ctrl[i], ctrl_delays[i]))
        ctrl_threads[i].start()
    #ends threads when we press ESC    
    while not (win32api.GetAsyncKeyState(win32con.VK_ESCAPE)):
        #sure i intended to have sample time = 0.5s
        #but i was missing deadlines
        s.enter(0.6, 1, ev_read_inputs[0].set)
        for i in range(0, 6):
            s.enter(0.6, 1, ev_ctrl[i].set)
            s.enter(1.2, 1, ev_ctrl[i].set)
            s.enter(1.8, 1, ev_ctrl[i].set)
            s.enter(2.4, 1, ev_ctrl[i].set)
        s.enter(1.2, 1, ev_read_inputs[1].set)
        s.enter(1.8, 1, ev_read_inputs[2].set)
        s.enter(2.4, 1, ev_read_inputs[3].set)
        s.run()
    keep_alive = False
    #run threads one last time to kill them and write logs
    for i in range(0, 4):
        ev_read_inputs[i].set()    
    for i in range(0, 6):
        ev_ctrl[i].set()