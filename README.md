
# OBS-FTC-Scene_Switcher

A scene switcher for OBS using the FTCLive websocket API and the OBS-Websocket-Py package. 


## Features

- Customizable scene names
- Customizable OBS websocket and scene settings
- 2/3/4 field support
- Skips finals and practice matches
- Incredibly buggy and confusing logging



## Deployment

To deploy this project run

```bash
  pyinstaller --noconsole --onefile --windowed --icon=icon.ico script.py

```

## Development

Download the python file

Go to the project directory

```bash
cd OBS-FTC-Scene_Switcher
```

Install dependencies

```bash
  pip install -r requirements.txt
```

Run the program

```bash
  python FTC_Switcher.py
```


## Roadmap

~~- Fix the YouTube description CSV~~

~~- Fix the YouTube description TXT~~
- Fix preview scene issues

~~- Add a GUI~~

- More and better logging

- Move away from python


## Acknowledgements

 - Ryan FTA 
   - Thanks for giving me your terrible powershell script and letting me make it better :)

