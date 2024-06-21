#!/bin/bash

# Loop through all directories in the current directory
for directory in */; do
  # Get the directory name without the trailing slash
  dir_name="${directory%*/}"

  # Skip the "venv" directory
  if [[ "$dir_name" == "venv" ]]; then
    continue
  fi

  # Create a zip archive named after the directory
  zip -r "$dir_name.cbz" "$directory"

  # Print a confirmation message
  echo "Zipped directory: $dir_name"
done