###########################################################
# AUTHOR  : Marius @ Hican - http://www.hican.nl - @hicannl
# DATE    : 23-04-2012
# COMMENT : This script renames all .jpg files to the
#           name of the .jpg parent folder recursively,
#           extended with an increasing number.
#           Put the script in the root of the folders'
#           parent folder.
#           NOTE: This script can also be used for renaming
#           other / multiple files, just adjust the filter!
###########################################################
$path = Split-Path -parent $MyInvocation.MyCommand.Definition

Function renameFiles
{
  # Loop through all directories
  $dirs = dir $path -Recurse | Where { $_.psIsContainer -eq $true }
  Foreach ($dir In $dirs)
  {
    # Set default value for addition to file name
    $i = 1
    $newdir = $dir.name + "_"
	# Search for the files set in the filter (*.jpg in this case)
    $files = Get-ChildItem -Path $dir.fullname -Filter *.jpg -Recurse
    Foreach ($file In $files)
    {
      # Check if a file exists
      If ($file)
      {
        # Split the name and rename it to the parent folder
        $split    = $file.name.split(".jpg")
        $replace  = $split[0] -Replace $split[0],($newdir + $i + ".jpg")

        # Trim spaces and rename the file
        $image_string = $file.fullname.ToString().Trim()
        "$split[0] renamed to $replace"
        Rename-Item "$image_string" "$replace"
	    $i++
      }
    }
  }
}
# RUN SCRIPT
renameFiles
"SCRIPT FINISHED"
