# CAD-coordinate-tool

A field-proven (2+ years of operational use) Python tool to import, export and save XYZ coordinates from CAD software. It's main use is to create lists of coordinates from CAD drawings and labels on the drawing itself.

## Description
This tool works with AutoCAD and some of its clones (tested on GstarCAD and nanoCAD).
It can export the coordinates of various objects (points, circles, lines, polylines, 
3D polylines, blocks, text) and reintroduce them in the CAD modelspace or save them 
to a .txt file.
It can also open a tab-separated coordinate file and import it into the CAD modelspace.
The text files it produces are compatible with various measuring instruments 
like Leica total stations.

## Requirements
- Python 3.x
- pywin32 (`pip install pywin32`)
- pandas (`pip install pandas`)
- numpy (`pip install numpy`)
- A running instance of AutoCAD, GstarCAD or nanoCAD

## Usage
1. Open your CAD software and load a drawing
2. Select the objects you want to export
3. Run the script: `python coordinate_tool.py`
4. Use the GUI buttons to export, import or save coordinates

## Supported objects
- Points
- Circles
- Lines
- Polylines (2D and 3D)
- Blocks / Block references
- Text and MText

## Screenshots
![GUI](CAD%20coordinates%20manager.jpg)
![CAD](CAD%20coordinates%20manager-NanoCAD.jpg)
