import random
import tkinter
import cv2
import threading
import time
import os
import win32com.client
import win32api
import pythoncom
from datetime import datetime
import sys
import shutil
import pandas as pd

camwidth = 640
camheight = 480
toggle = 1;
recordingIndex = -999
cam_id = 0
exp_init = "(blank)"

num_runs = 3
min_iti = 6
max_iti = 10

fish_id = "Z1"
sex = "(blank)"
genotype = "(blank)"
notes = "(blank)"

cue_t = 6
tone_dur = 4
pre_stim_t = 4
pre_rew_av_t = 4
rew_av_t = 4
post_rew_av_t = 5

filename = r'C:\Users\Kanwal\PycharmProjects\zebradata\paradigms.pptx'

tonePlaying = 0
videoPlaying = 0

forbidden = ["/", "<", ">", ":", '"', "\\", "|", "?", "*",
             chr(0), chr(1), chr(2), chr(3), chr(4), chr(5), chr(6), chr(7),
             chr(8), chr(9), chr(10), chr(11), chr(12), chr(13), chr(14), chr(15),
             chr(16), chr(17), chr(18), chr(19), chr(20), chr(21), chr(22), chr(23),
             chr(24), chr(25), chr(26), chr(27), chr(28), chr(29), chr(30), chr(31)]

global video_files
video_files = []

global trial_data
trial_data = []

now = datetime.now()
nowstr = now.strftime("%Y-%m-%d %H:%M:%S %p")

class VideoRecorder():

    # Video class based on openCV
    def __init__(self, paradigm):

        self.open = True
        self.device_index = cam_id
        self.fps = 20  # fps should be the minimum constant  rate at which the camera can
        self.fourcc = "XVID"  # capture images (with no decrease in speed over time; testing is required)
        self.frameSize = (640, 480)  # video formats and sizes also depend and vary according to the camera used
        self.video_filename = str(fish_id) + "_" + str(datetime.now().strftime('%Y-%m-%d_%H.%M')) + "_" + str(paradigm) + ".avi"
        self.video_cap = cv2.VideoCapture(self.device_index)
        self.video_writer = cv2.VideoWriter_fourcc(*self.fourcc)
        self.video_out = cv2.VideoWriter(self.video_filename, self.video_writer, self.fps, self.frameSize)
        self.frame_counts = 1
        self.start_time = time.time()
        self.font = cv2.FONT_HERSHEY_PLAIN
        self.xy = 10, 10

    # Video starts being recorded
    def record(self):

        while (self.open == True):
            ret, video_frame = self.video_cap.read()
            if (tonePlaying == 1):
                cv2.putText(video_frame, str(datetime.now()), (20, 40),
                            self.font, 2, (0, 0, 0), 2, cv2.LINE_AA)
                cv2.circle(video_frame, (620, 20), 20, (0, 0, 255), -1)

            elif videoPlaying == 1:
                cv2.putText(video_frame, str(datetime.now()), (20, 40),
                            self.font, 2, (0, 0, 0), 2, cv2.LINE_AA)
                cv2.rectangle(video_frame, (580, 10), (620, 40), (255, 0, 0), -1)

            else:
                cv2.putText(video_frame, str(datetime.now()), (20, 40),
                            self.font, 2, (255, 255, 255), 2, cv2.LINE_AA)

            if (ret == True):

                self.video_out.write(video_frame)
                self.frame_counts += 1
                time.sleep(0.05)
                gray = cv2.cvtColor(video_frame, cv2.COLOR_BGR2GRAY)
                cv2.imshow('video_frame', gray)
                cv2.waitKey(1)

            else:
                break

    def markerOn(self):
        cv2.drawMarker()

    # Finishes the video recording therefore the thread too
    def stop(self):

        if self.open == True:

            self.open = False
            self.video_out.release()
            self.video_cap.release()
            cv2.destroyAllWindows()

        else:
            pass

    # Launches the video recording function using a thread
    def start(self):
        video_thread = threading.Thread(target=self.record)
        video_thread.start()


def start_PPTrecording(filename):

    paradigm_slides = [['cf', 12], ['dfm', 7], ['ufm', 2]]
    all_runs = [['cf', 0], ['dfm', 0], ['ufm', 0]]

    fixed_times = [1, cue_t*1000, 1, 1]
    pythoncom.CoInitialize()
    app = win32com.client.Dispatch("PowerPoint.Application")
    app.Visible = 1
    app.Presentations.Open(FileName=filename)
    app.ActivePresentation.SlideShowSettings.Run()

    nov_test_len = 60
    print("Trial Onset:", nowstr)
    fixed = sum(fixed_times) / 1000
    vars = fixed + pre_stim_t + tone_dur + pre_rew_av_t + rew_av_t + post_rew_av_t
    if len(sys.argv) > 1:
        if sys.argv[1] == 'novel':
            vars += nov_test_len
    trial_min = (vars + min_iti) * num_runs
    trial_max = (vars + max_iti) * num_runs
    print("Length of Trial:", round((trial_min / 60), 1), "-", round((trial_max / 60), 1), "minutes")

    # 6-min novel environment test
    if len(sys.argv) > 1:
        if sys.argv[1] == 'novel':
            today = datetime.today()
            run_date = today.strftime('%m-%d-%y')
            run_time = today.strftime('%H.%M.%S')
            novtest_vthread = VideoRecorder('novelenvtest')
            novtest_data = [novtest_vthread.video_filename, fish_id, sex, genotype, run_date, run_time, exp_init,
                            'na', 'na', 'na', cam_id, notes, 'na', 'na', 'na', 'na', 'na', 'na', nov_test_len]
            trial_data.append(novtest_data)
            video_files.append(novtest_vthread.video_filename)
            app.SlideShowWindows(1).View.GotoSlide(1)
            novtest_vthread.start()
            print("novel environment test")
            time.sleep(nov_test_len)  # change to 360 for true trials
            novtest_vthread.stop()

    # loop through paradigm presentations and record from pre_stim_t to post rew_av_t
    for i in range(num_runs):
        this_run = random.choice(paradigm_slides)
        iti = random.randint(min_iti, max_iti)
        today = datetime.today()
        run_date = today.strftime('%m-%d-%y')
        run_time = today.strftime('%H.%M.%S')

        print('run', i + 1, ':', this_run[0], 'ITI:', iti, "onset:", run_time)

        video_thread = VideoRecorder(this_run[0])
        video_files.append(video_thread.video_filename)
        run_data = [video_thread.video_filename, fish_id, sex, genotype, run_date, run_time, exp_init,
                                     num_runs, min_iti, max_iti, cam_id, notes,
                                     pre_stim_t, cue_t, tone_dur, pre_rew_av_t, rew_av_t, post_rew_av_t, 'na']
        trial_data.append(run_data)

        video_thread.start()

        win32api.Sleep((pre_rew_av_t * 1000) + 2000)  # pre-stimulus time
        app.SlideShowWindows(1).View.GotoSlide(this_run[1])  # advance to screen cue
        win32api.Sleep(fixed_times[0])  # fixed 1
        app.SlideShowWindows(1).View.Next()  # play screen cue
        win32api.Sleep(fixed_times[1])  # fixed 2
        app.SlideShowWindows(1).View.Next()  # advance to sound slide
        win32api.Sleep(fixed_times[2])  # fixed 3
        app.SlideShowWindows(1).View.Next()  # play CF/FM
        global tonePlaying
        tonePlaying = 1
        win32api.Sleep(tone_dur * 1000)  # Analysis Period
        tonePlaying = 0
        app.SlideShowWindows(1).View.Next()  # advance to black slide
        win32api.Sleep(pre_stim_t * 1000)  # pre_rew_av_t interval
        app.SlideShowWindows(1).View.Next()  # advance to video slide
        win32api.Sleep(fixed_times[3])  # fixed 5
        app.SlideShowWindows(1).View.Next()  # start video
        global videoPlaying
        videoPlaying = 1
        win32api.Sleep(rew_av_t * 1000)  # rew_av_t time
        videoPlaying = 0
        app.SlideShowWindows(1).View.Next()  # advance to black slide

        for y, j in enumerate(all_runs):
            if this_run[0] == j[0]:
                j[1] += 1
            if j[1] == num_runs / 3:
                paradigm_slides.pop(y)
                all_runs.pop(y)

        time.sleep(post_rew_av_t)
        video_thread.stop()
        time.sleep(iti - post_rew_av_t)

        if len(all_runs) == 0:
            app.SlideShowWindows(1).View.GotoSlide(1)
            pythoncom.CoUninitialize()
            print("Presentation finished.")
            break

def main_():
    start_PPTrecording(filename)

    transcript_fn = 'transcript' + str(datetime.now().strftime('%Y-%m-%d_%H.%M')) + '.csv'
    trial_df = pd.DataFrame(trial_data, columns=['vid_file', 'fish_id', 'sex', 'genotype', 'date', 'time',
                                                 'exp_init','num_runs', 'min_iti', 'max_iti', 'cam_id', 'notes',
                                                 'pre_stim_t', 'cue_t', 'tone_dur', 'pre_rew_av_t', 'rew_av_t',
                                                 'post_rew_av_t', 'novtest_len'])
    trial_df.to_csv(transcript_fn)

    video_files.append(transcript_fn)
    directory = fish_id + "_trial_" + str(datetime.now().strftime('%Y-%m-%d_%H.%M'))
    parent_dir = "C:/Users/Kanwal/PycharmProjects/zebradata/"
    path = os.path.join(parent_dir, directory)
    os.mkdir(path)
    print("Directory '% s' created" % directory)

    for file in video_files:
        original = parent_dir + file
        target = path + "/" + file
        shutil.move(original, target)
    print("Trial files organized.")

    B.invoke()


def tkinter_start():
    top = tkinter.Tk()

    def action():
        global toggle
        toggle *= -1
        if (toggle == -1):
            print("Stopped Recording, exiting program")
            os._exit(0)
            cv2.destroyAllWindows()
    global B
    B = tkinter.Button(top, text="Stop Recording", command=action)

    B.pack()
    top.mainloop()


def supermain():
    t1 = threading.Thread(target=main_)
    t2 = threading.Thread(target=tkinter_start)
    t1.start()
    t2.start()
    t1.join()
    t2.join()


def startup():
    top1 = tkinter.Tk()
    top1.title("Zfish Interface")
    mainpanel = tkinter.PanedWindow(orient=tkinter.VERTICAL)
    mainpanel.pack(fill=tkinter.BOTH, expand=1)
    panel1 = tkinter.PanedWindow(mainpanel)
    panel1.pack(fill=tkinter.BOTH, expand=1)

    top1.geometry('360x440')

    def c():
        global fish_id
        if (not (txt7.get() == "")):
            fish_id = txt7.get()

        global sex
        if (not (txt8.get() == "")):
            sex = txt8.get()

        global genotype
        if (not (txt9.get() == "")):
            genotype = txt9.get()

        global exp_init
        if (not (txt3.get() == "")):
            exp_init = txt3.get()

        global num_runs
        if (not (txt4.get() == "")):
            num_runs = int(txt4.get())

        global min_iti
        if (not (txt5.get() == "")):
            min_iti = int(txt5.get())

        global max_iti
        if (not (txt51.get() == "")):
            max_iti = int(txt51.get())

        global cam_id
        if (not (txt1.get() == "")):
            cam_id = int(txt1.get())

        global notes
        if (not (txt10.get() == "")):
            notes = txt10.get()

        global pre_stim_t
        if (not (txt52.get() == "")):
            pre_stim_t = int(txt52.get())

        global tone_dur
        if (not (txt101.get() == "")):
            tone_dur = int(txt101.get())

        global pre_rew_av_t
        if (not (txt53.get() == "")):
            pre_rew_av_t = int(txt53.get())

        global rew_av_t
        if (not (txt54.get() == "")):
            rew_av_t = int(txt54.get())

        global post_rew_av_t
        if (not (txt100.get() == "")):
            post_rew_av_t = int(txt100.get())

        top1.destroy()

        return 1

    def val(char):
        if str.isdigit(char) or char == "":
            return True
        else:
            return False

    def valfn(char):
        if char in forbidden:
            return False
        else:
            return True

    val2 = (top1.register(val))
    valfn2 = (top1.register(valfn))

    panel2 = tkinter.PanedWindow(mainpanel, orient=tkinter.VERTICAL)
    panel2.pack()

    panel1 = tkinter.PanedWindow(panel2, orient=tkinter.HORIZONTAL)
    panel1.pack(anchor="w")
    label1 = tkinter.Label(top1, text="Cam Id: ", anchor="w", font=("Arial", 12))
    panel1.add(label1)
    txt1 = tkinter.Entry(top1, validate='all', validatecommand=(val2, '%P'))
    panel1.add(txt1)

    panel3 = tkinter.PanedWindow(panel2, orient=tkinter.HORIZONTAL)
    panel3.pack(anchor="w")
    label3 = tkinter.Label(top1, text="User Initials: ", anchor="w", font=("Arial", 12))
    panel3.add(label3)
    txt3 = tkinter.Entry(top1, validate='all')
    panel3.add(txt3)

    panel4 = tkinter.PanedWindow(panel2, orient=tkinter.HORIZONTAL)
    panel4.pack(anchor="w")
    label4 = tkinter.Label(top1, text="Number of Runs/Recordings: ", anchor="w", font=("Arial", 12))
    panel4.add(label4)
    txt4 = tkinter.Entry(top1, validate='all', validatecommand=(val2, '%P'))
    panel4.add(txt4)

    panel5 = tkinter.PanedWindow(panel2, orient=tkinter.HORIZONTAL)
    panel5.pack(anchor="w")
    label5 = tkinter.Label(top1, text="ITI min (sec): ", anchor="w", font=("Arial", 12))
    panel5.add(label5)
    txt5 = tkinter.Entry(top1, validate='all', validatecommand=(val2, '%P'))
    panel5.add(txt5)

    panel51 = tkinter.PanedWindow(panel2, orient=tkinter.HORIZONTAL)
    panel51.pack(anchor="w")
    label51 = tkinter.Label(top1, text="ITI max (sec): ", anchor="w", font=("Arial", 12))
    panel51.add(label51)
    txt51 = tkinter.Entry(top1, validate='all', validatecommand=(val2, '%P'))
    panel51.add(txt51)

    panel52 = tkinter.PanedWindow(panel2, orient=tkinter.HORIZONTAL)
    panel52.pack(anchor="w")
    label52 = tkinter.Label(top1, text="Pre-stimulus Time: ", anchor="w", font=("Arial", 12))
    panel52.add(label52)
    txt52 = tkinter.Entry(top1, validate='all', validatecommand=(val2, '%P'))
    panel52.add(txt52)

    panel53 = tkinter.PanedWindow(panel2, orient=tkinter.HORIZONTAL)
    panel53.pack(anchor="w")
    label53 = tkinter.Label(top1, text="Pre-reward Time: ", anchor="w", font=("Arial", 12))
    panel53.add(label53)
    txt53 = tkinter.Entry(top1, validate='all', validatecommand=(val2, '%P'))
    panel53.add(txt53)

    panel54 = tkinter.PanedWindow(panel2, orient=tkinter.HORIZONTAL)
    panel54.pack(anchor="w")
    label54 = tkinter.Label(top1, text="Reward/Aversion Time: ", anchor="w", font=("Arial", 12))
    panel54.add(label54)
    txt54 = tkinter.Entry(top1, validate='all', validatecommand=(val2, '%P'))
    panel54.add(txt54)

    panel100 = tkinter.PanedWindow(panel2, orient=tkinter.HORIZONTAL)
    panel100.pack(anchor="w")
    label100 = tkinter.Label(top1, text="Post-reward time: ", anchor="w", font=("Arial", 12))
    panel100.add(label100)
    txt100 = tkinter.Entry(top1, validate='all', validatecommand=(val2, '%P'))
    panel100.add(txt100)

    panel101 = tkinter.PanedWindow(panel2, orient=tkinter.HORIZONTAL)
    panel101.pack(anchor="w")
    label101 = tkinter.Label(top1, text="Tone Duration: ", anchor="w", font=("Arial", 12))
    panel101.add(label101)
    txt101 = tkinter.Entry(top1, validate='all', validatecommand=(val2, '%P'))
    panel101.add(txt101)

    panel6 = tkinter.PanedWindow(panel2, orient=tkinter.VERTICAL)
    panel6.pack(anchor="w")
    label6 = tkinter.Label(top1, text="Fish Information: ", anchor='center', font=("Arial", 12))
    label6.pack(anchor='center')
    label6.config(font=("Arial", 24))
    panel6.add(label6)

    panel7 = tkinter.PanedWindow(panel2, orient=tkinter.HORIZONTAL)
    panel7.pack(anchor="w")
    label7 = tkinter.Label(top1, text="Fish ID: ", anchor="w", font=("Arial", 12))
    panel7.add(label7)
    txt7 = tkinter.Entry(top1, validate='all', validatecommand=(valfn2, '%P'))
    panel7.add(txt7)

    panel8 = tkinter.PanedWindow(panel2, orient=tkinter.HORIZONTAL)
    panel8.pack(anchor="w")
    label8 = tkinter.Label(top1, text="sex (M/F): ", anchor="w", font=("Arial", 12))
    panel8.add(label8)
    txt8 = tkinter.Entry(top1, validate='all')
    panel8.add(txt8)

    panel9 = tkinter.PanedWindow(panel2, orient=tkinter.HORIZONTAL)
    panel9.pack(anchor="w")
    label9 = tkinter.Label(top1, text="Genotype (W/M/T): ", anchor="w", font=("Arial", 12))
    panel9.add(label9)
    txt9 = tkinter.Entry(top1, validate='all')
    panel9.add(txt9)

    panel10 = tkinter.PanedWindow(panel2, orient=tkinter.HORIZONTAL)
    panel10.pack(anchor="w")
    label10 = tkinter.Label(top1, text="Notes: ", anchor="w", font=("Arial", 12))
    panel10.add(label10)
    txt10 = tkinter.Entry(top1, validate='all')
    panel10.add(txt10)

    C = tkinter.Button(top1, text="GO", command=c)
    C.pack(side=tkinter.BOTTOM)
    top1.mainloop()


startup()
supermain()



