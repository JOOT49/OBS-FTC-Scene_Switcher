#!/usr/bin/env python3
"""
Direct conversion of FTC_OBS_SceneSwitches_Main.ps1 and FTC_OBS_SceneSwitches_v2_RUNME.ps1 into Python,
using obs-websocket-py for OBS communication.
DO NOT change how anything is called.
"""

import os
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


# Import obs-websocket-py library (version 4)
from obswebsocket import obsws, requests as obsrequests

#==================================================
#               Variables
#       (Set these directly in the script)
#==================================================

OBS_SERVERNAME = "192.168.1.154"
OBS_WEBSOCKET_PASSWORD = "secret"
OBS_WEBSOCKET_PORT = 4455  # Use an integer port

OBS_SCENENAME_FIELD2 = "Red Only (Field 2)"
OBS_SCENENAME_FIELD1 = "Blue Only (Field 1)"
# If these are not set, the script proceeds as a 2-field event.
OBS_SCENENAME_FIELD3 = "Field 3"
OBS_SCENENAME_FIELD4 = "Field 4"

FTCSERVER_NAME = "localhost"
FTCSERVER_EVENTCODE = "usjdb"

#==================================================
#               Helper Functions
#==================================================

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

#--------------------------------------------------
# OBS functions using obs-websocket-py (synchronous)
#--------------------------------------------------

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
        # Call GetCurrentScene() and use getSceneName() to extract the scene name
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
    """
    try:
        response = connection.call(obsrequests.GetStreamStatus())
        # If the response does not include outputDuration, default to 0
        return {"outputDuration": getattr(response, 'outputDuration', 0)}
    except Exception as e:
        WriteLog(f"Error getting OBS stream status: {e}")
        return {"outputDuration": 0}

#==================================================
#       FTC Websocket Client (Scorekeeper)
#==================================================

# Create concurrent queues for receiving and sending messages.
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

#==================================================
#               Main Script
#==================================================

def main():
    # Validate variables (check for OBS_SCENENAME_FIELD3/4)
    if OBS_SCENENAME_FIELD3 is None:
        errorMessage = "[ERROR] OBS_SCENENAME_FIELD3 is NOT set! Proceeding as 2 field event."
        WriteLog(errorMessage)
        Write_Host(errorMessage)
    if (OBS_SCENENAME_FIELD3 is not None) and (OBS_SCENENAME_FIELD4 is None):
        errorMessage = "[ERROR] OBS_SCENENAME_FIELD4 is NOT set! Proceeding as 3 field event."
        WriteLog(errorMessage)
        Write_Host(errorMessage)

    # Test connections
    if not test_connection(OBS_SERVERNAME):
        errorMessage = "[ERROR] Unable to connect to OBS, please check the IP or your network connection."
        WriteLog(errorMessage)
        Write_Host(errorMessage)
    if not test_connection(FTCSERVER_NAME):
        errorMessage = "[ERROR] Unable to connect to the FTC Scoring system. Please check the IP or your network connection."
        WriteLog(errorMessage)
        Write_Host(errorMessage)

    WriteLog("Code is Starting")

    # Connect to OBS
    try:
        obs_connection = Connect_OBS()
    except Exception as e:
        WriteLog(str(e))
        return

    # Connect to FTC WebSocket
    ftc_ws_url = f"ws://{FTCSERVER_NAME}/api/v2/stream/?code={FTCSERVER_EVENTCODE}"
    Write_Host("Connecting to FTC websocket...")
    try:
        ftc_ws = websocket.create_connection(ftc_ws_url)
        Write_Host("Connected to FTC websocket!")
    except Exception as e:
        WriteLog(f"Failed to connect to FTC WebSocket: {e}")
        return

    # Start threads for receiving and sending FTC messages
    recv_thread_obj = threading.Thread(target=ftc_recv_job, args=(ftc_ws,), daemon=True)
    send_thread_obj = threading.Thread(target=ftc_send_job, args=(ftc_ws,), daemon=True)
    recv_thread_obj.start()
    send_thread_obj.start()

    WriteLog("Code is Running")

    try:
        # Main processing loop – run indefinitely until interrupted.
        while True:
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

            # Debug: Log entire message payload.
            WriteLog(f"Received message: {message_obj}")

            updateType = message_obj.get("updateType")
            payload = message_obj.get("payload", {})
            shortName = payload.get("shortName", "")
            field = str(payload.get("field", ""))

            # Debug: Log extracted values.
            WriteLog(f"updateType: {updateType}, shortName: {shortName}, field: {field}")

            # Process messages based on updateType
            if updateType in ["SHOW_PREVIEW", "SHOW_MATCH"]:
                if not shortName.startswith("F-"):  # Skip finals
                    if field == "1":
                        current_scene = Get_OBSCurrentProgramScene(obs_connection)
                        if current_scene != OBS_SCENENAME_FIELD1:
                            Write_Host(f"[{time.strftime('%Y%m%d %H:%M:%S')}] Switching to FIELD {field}")
                            Set_OBSCurrentProgramScene(obs_connection, OBS_SCENENAME_FIELD1)
                            # Allow a short delay before setting preview scene.
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
                if not shortName.startswith("T-"):  # Qualification matches only
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
                        fieldnames = ["TimeStamp", "MatchName", "Red1", "Red2", "Blue1", "Blue2", "RedFinal", "BlueFinal"]
                        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                        if not file_exists:
                            writer.writeheader()
                        writer.writerow(new_row)
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
                            fieldnames = ["TimeStamp", "MatchName", "Red1", "Red2", "Blue1", "Blue2", "RedFinal", "BlueFinal"]
                            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                            writer.writeheader()
                            writer.writerows(rows)

            time.sleep(0.01)

    except KeyboardInterrupt:
        WriteLog("Code is stopping")
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

if __name__ == '__main__':
    main()
