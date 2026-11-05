# Video-to-Powerpoint
I (ChatGPT) made a python script that allows to convert a video into a powerpoint presentation

The script supports framerates from 0.1 to 60fps. Every frame per second is its own slide.


# Installation:
1. Download this repository or just video_to_pptx.py and requirements.txt
2. Then open terminal and go to the path of the folder the video_to_pptx.py file is located in
3. When you are in the correct folder run "py -m pip install requirements.txt"

# Use
1. Go to the folder video_to_pptx.py is located
2. Add the video to the folder
3. Run python video_to_pptx.py "(Name of the video).mp4" and add your arguments here

# List of all arguments:
1. Fps: add --fps (fps of the powerpoint)
2. Output: --output (target name of the presentation)
3. Max slides: --max-slides (max counts of slides in the presentation, default if no input is given: 1000000000 (close to 193 days with 60 slides per second))
4. Auto advance: --auto-advance (currently does nothing)
5. Advance seconds: --advance-seconds (currently does nothing)

An example would be:

  python video_to_pptx.py "My-Video.mp4" --fps 60 --output "My Video Presentation" --max-slides 600

