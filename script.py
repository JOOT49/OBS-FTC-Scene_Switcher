#!/usr/bin/env python3

import os
import sys
import time
import threading
import json
import csv
import re
import uuid
import subprocess
import requests
import queue
import websocket
import tkinter as tk
from tkinter import ttk

# Global exit flag
exit_requested = False

# ==================================================
#               Default Variables
#       (These will be updated by the GUI)
# ==================================================

OBS_SERVERNAME = ""
OBS_WEBSOCKET_PASSWORD = ""
OBS_WEBSOCKET_PORT = 4455  # Use an integer port

OBS_SCENENAME_FIELD2 = ""
OBS_SCENENAME_FIELD1 = ""
# If these are not set, the script proceeds as a 2-field event.
OBS_SCENENAME_FIELD3 = ""
OBS_SCENENAME_FIELD4 = ""

FTCSERVER_NAME = ""
FTCSERVER_EVENTCODE = ""


# ==================================================
#               Helper Functions
# ==================================================

def WriteLog(LogString=""):
    print(f"{time.strftime('%Y%m%d %H:%M:%S')} {LogString}")


def Write_Host(message):
    print(message)


def test_connection(host, count=5):
    """
    Simulates Test-Connection by pinging the host.
    """
    try:
        param = '-n' if os.name == 'nt' else '-c'
        command = ['ping', param, str(count), host]
        result = subprocess.run(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        return result.returncode == 0
    except Exception:
        return False


# --------------------------------------------------
# OBS functions using obs-websocket-py (synchronous)
# --------------------------------------------------

def Connect_OBS():
    """
    Connect to OBS using obs-websocket-py.
    """
    try:
        port = OBS_WEBSOCKET_PORT if OBS_WEBSOCKET_PORT else 4455
        connection = obsws(host=OBS_SERVERNAME, port=int(port), password=OBS_WEBSOCKET_PASSWORD)
        connection.connect()
        WriteLog("Connected to OBS!")
        return connection
    except Exception as e:
        raise Exception(f"Failed to connect to OBS: {e}")


def Get_OBSCurrentProgramScene(connection):
    """
    Gets the current scene using GetCurrentScene().
    """
    try:
        response = connection.call(obsrequests.GetCurrentScene())
        scene = response.getSceneName()  # Correct method to get the scene name
        WriteLog(f"OBS Current Scene: '{scene}'")
        return scene
    except Exception as e:
        WriteLog(f"Error getting current scene: {e}")
        return ""


def Set_OBSCurrentProgramScene(connection, SceneName):
    """
    Sets the current scene using SetCurrentScene.
    """
    try:
        WriteLog(f"Setting OBS current scene to: '{SceneName}'")
        connection.call(obsrequests.SetCurrentProgramScene(sceneName=SceneName))
    except Exception as e:
        WriteLog(f"Error setting current scene to {SceneName}: {e}")


def Set_OBSCurrentPreviewScene(connection, SceneName):
    """
    Sets the preview scene using SetCurrentPreviewScene.
    (Note: This only works if OBS is in Studio Mode.)
    """
    try:
        WriteLog(f"Setting OBS preview scene to: '{SceneName}'")
        connection.call(obsrequests.SetCurrentPreviewScene(sceneName=SceneName))
    except Exception as e:
        WriteLog(f"Error setting preview scene to {SceneName}: {e}")


def Get_OBSStreamStatus(connection):
    """
    Gets the OBS stream status.
    Returns a dictionary with 'outputDuration' in milliseconds.

    Updated: Uses the getter method getOutputDuration() (if available) to retrieve the stream duration.
    """
    try:
        response = connection.call(obsrequests.GetStreamStatus())
        duration = response.getOutputDuration() if hasattr(response, 'getOutputDuration') else 0
        return {"outputDuration": duration}
    except Exception as e:
        WriteLog(f"Error getting OBS stream status: {e}")
        return {"outputDuration": 0}


def generate_youtube_description():
    """
    Reads the CSV file and generates a YouTube description in a TXT file.
    """
    csv_file = f"{FTCSERVER_EVENTCODE}_YouTube_Description.csv"
    txt_file = f"{FTCSERVER_EVENTCODE}_YouTube_Description.txt"
    if not os.path.isfile(csv_file):
        return
    with open(csv_file, 'r', newline='') as csvfile:
        reader = csv.DictReader(csvfile)
        rows = list(reader)
    with open(txt_file, 'w') as txtfile:
        for row in rows:
            txtfile.write(f"Match: {row.get('MatchName', '')}\n")
            txtfile.write(f"Time: {row.get('TimeStamp', '')}\n")
            txtfile.write(
                f"Red Alliance: {row.get('Red1', '')} & {row.get('Red2', '')} - Score: {row.get('RedFinal', '')}\n")
            txtfile.write(
                f"Blue Alliance: {row.get('Blue1', '')} & {row.get('Blue2', '')} - Score: {row.get('BlueFinal', '')}\n")
            txtfile.write("-----------------------------\n")


# ==================================================
#       FTC Websocket Client (Scorekeeper)
# ==================================================

from obswebsocket import obsws, requests as obsrequests

recv_queue = queue.Queue()
send_queue = queue.Queue()
client_id = str(uuid.uuid4())


def ftc_recv_job(ws):
    """
    This function runs in a thread and continually receives messages from the FTC websocket.
    """
    while True:
        try:
            jsonResult = ws.recv()
            if jsonResult:
                recv_queue.put(jsonResult)
        except Exception:
            break
        time.sleep(0.01)


def ftc_send_job(ws):
    """
    This function runs in a thread and sends messages queued in send_queue.
    """
    while True:
        try:
            workitem = send_queue.get(timeout=1)
            ws.send(workitem)
        except queue.Empty:
            continue
        except Exception:
            break


# ==================================================
#               Main Script
# ==================================================

def main():
    global exit_requested

    if OBS_SCENENAME_FIELD3 is None:
        errorMessage = "[ERROR] OBS_SCENENAME_FIELD3 is NOT set! Proceeding as 2 field event."
        WriteLog(errorMessage)
        Write_Host(errorMessage)
    if (OBS_SCENENAME_FIELD3 is not None) and (OBS_SCENENAME_FIELD4 is None):
        errorMessage = "[ERROR] OBS_SCENENAME_FIELD4 is NOT set! Proceeding as 3 field event."
        WriteLog(errorMessage)
        Write_Host(errorMessage)

    if not test_connection(OBS_SERVERNAME):
        errorMessage = "[ERROR] Unable to connect to OBS, please check the IP or your network connection."
        WriteLog(errorMessage)
        Write_Host(errorMessage)
    if not test_connection(FTCSERVER_NAME):
        errorMessage = "[ERROR] Unable to connect to the FTC Scoring system. Please check the IP or your network connection."
        WriteLog(errorMessage)
        Write_Host(errorMessage)

    WriteLog("Code is Starting")

    try:
        obs_connection = Connect_OBS()
    except Exception as e:
        WriteLog(str(e))
        return

    ftc_ws_url = f"ws://{FTCSERVER_NAME}/api/v2/stream/?code={FTCSERVER_EVENTCODE}"
    Write_Host("Connecting to FTC websocket...")
    try:
        ftc_ws = websocket.create_connection(ftc_ws_url)
        Write_Host("Connected to FTC websocket!")
    except Exception as e:
        WriteLog(f"Failed to connect to FTC WebSocket: {e}")
        return

    recv_thread_obj = threading.Thread(target=ftc_recv_job, args=(ftc_ws,), daemon=True)
    send_thread_obj = threading.Thread(target=ftc_send_job, args=(ftc_ws,), daemon=True)
    recv_thread_obj.start()
    send_thread_obj.start()

    WriteLog("Code is Running")

    try:
        while True:
            if exit_requested:
                WriteLog("Exit requested; breaking main loop.")
                break

            try:
                msg = recv_queue.get(timeout=1)
            except queue.Empty:
                time.sleep(0.1)
                continue

            if msg == "pong":
                continue

            try:
                message_obj = json.loads(msg)
            except Exception:
                WriteLog(f"Error converting message to JSON: {msg}")
                continue

            WriteLog(f"Received message: {message_obj}")

            updateType = message_obj.get("updateType")
            payload = message_obj.get("payload", {})
            shortName = payload.get("shortName", "")
            field = str(payload.get("field", ""))

            WriteLog(f"updateType: {updateType}, shortName: {shortName}, field: {field}")

            if updateType in ["SHOW_PREVIEW", "SHOW_MATCH"]:
                if not shortName.startswith("F-"):
                    if field == "1":
                        current_scene = Get_OBSCurrentProgramScene(obs_connection)
                        if current_scene != OBS_SCENENAME_FIELD1:
                            Write_Host(f"[{time.strftime('%Y%m%d %H:%M:%S')}] Switching to FIELD {field}")
                            Set_OBSCurrentProgramScene(obs_connection, OBS_SCENENAME_FIELD1)
                            time.sleep(0.2)
                            Set_OBSCurrentPreviewScene(obs_connection, OBS_SCENENAME_FIELD2)
                        else:
                            Write_Host(f"[{time.strftime('%Y%m%d %H:%M:%S')}] Already on FIELD {field}")
                    elif field == "2":
                        current_scene = Get_OBSCurrentProgramScene(obs_connection)
                        if current_scene != OBS_SCENENAME_FIELD2:
                            Write_Host(f"[{time.strftime('%Y%m%d %H:%M:%S')}] Switching to FIELD {field}")
                            Set_OBSCurrentProgramScene(obs_connection, OBS_SCENENAME_FIELD2)
                            time.sleep(0.2)
                            if OBS_SCENENAME_FIELD3:
                                Set_OBSCurrentPreviewScene(obs_connection, OBS_SCENENAME_FIELD3)
                            else:
                                Set_OBSCurrentPreviewScene(obs_connection, OBS_SCENENAME_FIELD1)
                        else:
                            Write_Host(f"[{time.strftime('%Y%m%d %H:%M:%S')}] Already on FIELD {field}")
                    elif field == "3":
                        current_scene = Get_OBSCurrentProgramScene(obs_connection)
                        if current_scene != OBS_SCENENAME_FIELD3:
                            Write_Host(f"[{time.strftime('%Y%m%d %H:%M:%S')}] Switching to FIELD {field}")
                            Set_OBSCurrentProgramScene(obs_connection, OBS_SCENENAME_FIELD3)
                            time.sleep(0.2)
                            if OBS_SCENENAME_FIELD4:
                                Set_OBSCurrentPreviewScene(obs_connection, OBS_SCENENAME_FIELD4)
                            else:
                                Set_OBSCurrentPreviewScene(obs_connection, OBS_SCENENAME_FIELD1)
                        else:
                            Write_Host(f"[{time.strftime('%Y%m%d %H:%M:%S')}] Already on FIELD {field}")
                    elif field == "4":
                        current_scene = Get_OBSCurrentProgramScene(obs_connection)
                        if current_scene != OBS_SCENENAME_FIELD4:
                            Write_Host(f"[{time.strftime('%Y%m%d %H:%M:%S')}] Switching to FIELD {field}")
                            Set_OBSCurrentProgramScene(obs_connection, OBS_SCENENAME_FIELD4)
                            time.sleep(0.2)
                            Set_OBSCurrentPreviewScene(obs_connection, OBS_SCENENAME_FIELD1)
                        else:
                            Write_Host(f"[{time.strftime('%Y%m%d %H:%M:%S')}] Already on FIELD {field}")
            elif updateType == "MATCH_START":
                if not shortName.startswith("T-"):
                    stream_status = Get_OBSStreamStatus(obs_connection)
                    try:
                        ms = stream_status.get("outputDuration", 0)
                        seconds = ms / 1000.0
                        TimeStamp = time.strftime("%H:%M:%S", time.gmtime(seconds))
                    except Exception:
                        TimeStamp = "00:00:00"
                    new_row = {
                        "TimeStamp": TimeStamp,
                        "MatchName": shortName,
                        "Red1": "",
                        "Red2": "",
                        "Blue1": "",
                        "Blue2": "",
                        "RedFinal": "",
                        "BlueFinal": ""
                    }
                    csv_file = f"{FTCSERVER_EVENTCODE}_YouTube_Description.csv"
                    file_exists = os.path.isfile(csv_file)
                    with open(csv_file, 'a', newline='') as csvfile:
                        fieldnames = ["TimeStamp", "MatchName", "Red1", "Red2", "Blue1", "Blue2", "RedFinal",
                                      "BlueFinal"]
                        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                        if not file_exists:
                            writer.writeheader()
                        writer.writerow(new_row)
                    generate_youtube_description()
            elif updateType == "MATCH_COMMIT":
                if shortName.startswith("Q"):
                    url = f"http://{FTCSERVER_NAME}/api/v1/events/{FTCSERVER_EVENTCODE}/matches/{payload.get('number')}/"
                    res = requests.get(url).json()
                    csv_file = f"{FTCSERVER_EVENTCODE}_YouTube_Description.csv"
                    rows = []
                    if os.path.isfile(csv_file):
                        with open(csv_file, 'r', newline='') as csvfile:
                            reader = csv.DictReader(csvfile)
                            for row in reader:
                                if row.get("MatchName") == shortName:
                                    row["Red1"] = res.get("matchBrief", {}).get("red", {}).get("team1", "")
                                    row["Red2"] = res.get("matchBrief", {}).get("red", {}).get("team2", "")
                                    row["Blue1"] = res.get("matchBrief", {}).get("blue", {}).get("team1", "")
                                    row["Blue2"] = res.get("matchBrief", {}).get("blue", {}).get("team2", "")
                                    row["RedFinal"] = res.get("redScore", "")
                                    row["BlueFinal"] = res.get("blueScore", "")
                                rows.append(row)
                        with open(csv_file, 'w', newline='') as csvfile:
                            fieldnames = ["TimeStamp", "MatchName", "Red1", "Red2", "Blue1", "Blue2", "RedFinal",
                                          "BlueFinal"]
                            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                            writer.writeheader()
                            writer.writerows(rows)
                    generate_youtube_description()

            time.sleep(0.01)

    except KeyboardInterrupt:
        WriteLog("Code is stopping (KeyboardInterrupt)")
    except Exception as e:
        WriteLog(f"Error in main loop: {e}")
    finally:
        Write_Host("Closing FTC WS connection")
        try:
            ftc_ws.close()
        except Exception:
            pass
        Write_Host("Disconnecting from OBS")
        try:
            obs_connection.disconnect()
        except Exception:
            pass


def launch_config_gui():
    """
    Opens a Tkinter GUI to configure the global variables.
    Displays the 'intodeep.png' logo (scaled down with preserved aspect ratio) at the top.
    """
    config = {}
    root = tk.Tk()
    root.title("Configuration Settings")

    # Create a frame for the logo and pack it at the top.
    logo_frame = tk.Frame(root)
    logo_frame.pack(side=tk.TOP, pady=10)
    try:
        from PIL import Image, ImageTk
        logo_image = Image.open("intodeep.png")
        try:
            resample_method = Image.Resampling.LANCZOS
        except AttributeError:
            resample_method = Image.LANCZOS
        # Use thumbnail to scale down preserving aspect ratio
        logo_image.thumbnail((128, 128), resample_method)
        logo_photo = ImageTk.PhotoImage(logo_image)
        logo_label = tk.Label(logo_frame, image=logo_photo)
        logo_label.image = logo_photo
        logo_label.pack()
    except Exception as e:
        WriteLog("Could not load logo image: " + str(e))

    # Create a frame for the form entries.
    form_frame = tk.Frame(root)
    form_frame.pack(side=tk.TOP, pady=10)

    labels = [
        ("OBS Server Name:", "OBS_SERVERNAME"),
        ("OBS Websocket Password:", "OBS_WEBSOCKET_PASSWORD"),
        ("OBS Websocket Port:", "OBS_WEBSOCKET_PORT"),
        ("OBS Scene Name Field 2:", "OBS_SCENENAME_FIELD2"),
        ("OBS Scene Name Field 1:", "OBS_SCENENAME_FIELD1"),
        ("OBS Scene Name Field 3:", "OBS_SCENENAME_FIELD3"),
        ("OBS Scene Name Field 4:", "OBS_SCENENAME_FIELD4"),
        ("FTC Server Name:", "FTCSERVER_NAME"),
        ("FTC Event Code:", "FTCSERVER_EVENTCODE")
    ]

    default_values = {
        "OBS_SERVERNAME": OBS_SERVERNAME,
        "OBS_WEBSOCKET_PASSWORD": OBS_WEBSOCKET_PASSWORD,
        "OBS_WEBSOCKET_PORT": str(OBS_WEBSOCKET_PORT),
        "OBS_SCENENAME_FIELD2": OBS_SCENENAME_FIELD2,
        "OBS_SCENENAME_FIELD1": OBS_SCENENAME_FIELD1,
        "OBS_SCENENAME_FIELD3": OBS_SCENENAME_FIELD3,
        "OBS_SCENENAME_FIELD4": OBS_SCENENAME_FIELD4,
        "FTCSERVER_NAME": FTCSERVER_NAME,
        "FTCSERVER_EVENTCODE": FTCSERVER_EVENTCODE
    }

    entries = {}
    for i, (label_text, key) in enumerate(labels):
        lbl = ttk.Label(form_frame, text=label_text)
        lbl.grid(row=i, column=0, padx=5, pady=5, sticky='e')
        entry = ttk.Entry(form_frame)
        entry.insert(0, default_values[key])
        entry.grid(row=i, column=1, padx=5, pady=5, sticky='w')
        entries[key] = entry

    def on_start():
        config["OBS_SERVERNAME"] = entries["OBS_SERVERNAME"].get()
        config["OBS_WEBSOCKET_PASSWORD"] = entries["OBS_WEBSOCKET_PASSWORD"].get()
        try:
            config["OBS_WEBSOCKET_PORT"] = int(entries["OBS_WEBSOCKET_PORT"].get())
        except ValueError:
            config["OBS_WEBSOCKET_PORT"] = 4455
        config["OBS_SCENENAME_FIELD2"] = entries["OBS_SCENENAME_FIELD2"].get()
        config["OBS_SCENENAME_FIELD1"] = entries["OBS_SCENENAME_FIELD1"].get()
        config["OBS_SCENENAME_FIELD3"] = entries["OBS_SCENENAME_FIELD3"].get()
        config["OBS_SCENENAME_FIELD4"] = entries["OBS_SCENENAME_FIELD4"].get()
        config["FTCSERVER_NAME"] = entries["FTCSERVER_NAME"].get()
        config["FTCSERVER_EVENTCODE"] = entries["FTCSERVER_EVENTCODE"].get()
        root.destroy()

    start_button = ttk.Button(root, text="Save Configurations", command=on_start)
    start_button.pack(pady=10)

    root.mainloop()
    return config


def launch_exit_window():
    """
    Opens a Tkinter window with the logo (scaled down with preserved aspect ratio),
    a label that says "FTC Scene switching is running", and a single button labeled "Exit".
    Clicking the button sets the exit flag.
    """
    global exit_requested
    exit_root = tk.Tk()
    exit_root.title("Exit Application")

    # Logo frame
    logo_frame = tk.Frame(exit_root)
    logo_frame.pack(side=tk.TOP, pady=10)
    try:
        from PIL import Image, ImageTk
        logo_image = Image.open("intodeep.png")
        try:
            resample_method = Image.Resampling.LANCZOS
        except AttributeError:
            resample_method = Image.LANCZOS
        logo_image.thumbnail((128, 128), resample_method)
        logo_photo = ImageTk.PhotoImage(logo_image)
        logo_label = tk.Label(logo_frame, image=logo_photo)
        logo_label.image = logo_photo
        logo_label.pack()
    except Exception as e:
        WriteLog("Could not load logo image in exit window: " + str(e))

    # Text label indicating status
    status_label = ttk.Label(exit_root, text="FTC Scene switching is running")
    status_label.pack(pady=5)

    # Exit button
    exit_button = ttk.Button(exit_root, text="Exit", command=lambda: on_exit(exit_root))
    exit_button.pack(padx=20, pady=20)

    exit_root.mainloop()


def on_exit(window):
    global exit_requested
    exit_requested = True
    window.destroy()


# ==================================================
#                    Entry Point
# ==================================================

if __name__ == '__main__':
    config = launch_config_gui()
    OBS_SERVERNAME = config.get("OBS_SERVERNAME", OBS_SERVERNAME)
    OBS_WEBSOCKET_PASSWORD = config.get("OBS_WEBSOCKET_PASSWORD", OBS_WEBSOCKET_PASSWORD)
    OBS_WEBSOCKET_PORT = config.get("OBS_WEBSOCKET_PORT", OBS_WEBSOCKET_PORT)
    OBS_SCENENAME_FIELD2 = config.get("OBS_SCENENAME_FIELD2", OBS_SCENENAME_FIELD2)
    OBS_SCENENAME_FIELD1 = config.get("OBS_SCENENAME_FIELD1", OBS_SCENENAME_FIELD1)
    OBS_SCENENAME_FIELD3 = config.get("OBS_SCENENAME_FIELD3", OBS_SCENENAME_FIELD3)
    OBS_SCENENAME_FIELD4 = config.get("OBS_SCENENAME_FIELD4", OBS_SCENENAME_FIELD4)
    FTCSERVER_NAME = config.get("FTCSERVER_NAME", FTCSERVER_NAME)
    FTCSERVER_EVENTCODE = config.get("FTCSERVER_EVENTCODE", FTCSERVER_EVENTCODE)

    main_thread = threading.Thread(target=main, daemon=True)
    main_thread.start()

    launch_exit_window()

    main_thread.join()
    sys.exit(0)
