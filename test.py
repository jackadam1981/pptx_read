# pip install moviepy
import moviepy.video.io.ImageSequenceClip

clip = moviepy.video.io.ImageSequenceClip.ImageSequenceClip(['221.JPG', ], fps=0.1)
clip.write_videofile('Movie1.mp4')
