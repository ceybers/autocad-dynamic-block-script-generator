# autocad-dynamic-block-script-generator
Excel VBA that generates an AutoCAD script file that inserts dynamic blocks in bulk.

## Background

We had a task where we needed to label 100+ pieces of equipment in an AutoCAD drawing with data stored in a spreadsheet. 

1. In AutoCAD LT, we create a dynamic block, with one attribute for each field we require. 
2. We place all the data in a table in an Excel spreadsheet.
3. We run a VBA subroutine that generates an AutoCAD script. This script contains all the commands one would type in the AutoCAD commandline to insert the dynamic blocks manually.
4. In AutoCAD, we use the `SCRIPT` command, and run the .scr file.

## AutoCAD Script Files?

The `SCRIPT` command will run each line in a script file in the commandline of the active drawing. The command we use to insert dynamic blocks is `-INSERT`. It should then prompt you for the below data. You should run it once manually, as your system variables might affect the order and/or prompts.

By default the `SCRIPT` command will try to look for a .scr file with the same filename as your drawing, so setting that will save you an extra step.

1. Name of your block
2. Coordinates to place the block at
3. Scale in the x-axis
4. Scale in the y-axis
5. Rotation
6. Text for the first attribute
7. Text for the second attribute (and so forth)

## xOffset and xStep?

If you don't have the coordinates of the equipment, it might be easier to just insert all the blocks in an array in an empty space on your drawing. 

Once you've moved them in place, you can use the LIST command and parse the output to get the coordinates of the blocks.
