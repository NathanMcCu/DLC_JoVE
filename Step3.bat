#!/bin/bash

# Input directory containing MJPEG videos
input_dir="C:\Users\Work\Desktop\Levi Comp\Raw_mjpeg"

# Output directory for MP4 videos
output_dir="C:\Users\Work\Desktop\Levi Comp\Converted_mp4"

# Set the target frame rate
frame_rate=10

# Check if the output directory exists; if not, create it
mkdir -p "$output_dir"

# Loop through MJPEG videos in the input directory and convert each one
for video in "$input_dir"/*.mjpeg; do
    if [ -e "$video" ]; then
        # Extract the video file name without the path
        file_name=$(basename "$video")
        file_name_no_ext="${file_name%.*}"

        # Define the output file path using the same name as the input file
        output_file="$output_dir/$file_name_no_ext.mp4"

        # Run FFmpeg command to convert the MJPEG video to MP4
        ffmpeg -r 10 -i  -c:v libx264 -pix_fmt yuv420p -vf "hflip,vflip" -r 10 

        echo "Converted $video to $output_file"
    fi
done
