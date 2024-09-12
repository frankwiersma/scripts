# Transparent Timer GIFs for PowerPoint

This project generates transparent timer GIFs that can be used in PowerPoint presentations. Each GIF represents a one-minute countdown, allowing for skippable timers across slides.

## Installation

1. Ensure you have Python 3.x installed on your system.

2. Install the required dependencies:

   ```
   pip install Wand
   ```

3. Install ImageMagick:
   - **Windows**: Download and install from [ImageMagick Website](https://imagemagick.org/script/download.php)
   - **Mac**: Use Homebrew: `brew install imagemagick`
   - **Linux**: Use your package manager, e.g., `sudo apt-get install imagemagick`

## Usage

1. Run the Python script to generate the timer GIFs:

   ```
   python TimerGifCreator.py
   ```

2. Find the generated GIFs in the `timer_gifs` directory.

3. In PowerPoint:
   - Insert a new slide for each minute of your timer.
   - On each slide, insert the corresponding GIF (Insert > Picture > [select GIF]).
   - Position the GIF as desired on each slide.
   - Set slide transition to "After: 1:00" for automatic progression, or advance manually for a skippable timer.

## Tips

- The transparent background allows the timer to overlay other content on your slides.
- For manual progression, you can skip to any point in the countdown by navigating through the slides.
- Adjust slide backgrounds or add content as needed; the timer will remain on top.

Enjoy your customizable, transparent timers in PowerPoint!
