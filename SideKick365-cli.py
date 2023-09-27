# Command line interface for SideKick365
# Steve Goodman 2023/09/27
# allows the following commands:
# --wordtopowerpoint --wordfile <wordfile> --powerpointfile <powerpointfile>

import argparse
import os
import sys
import sidekick365

# check a valid command (--wordtopowerpoint) was specified on the command line was specifed and if not print help
# create a parser
parser = argparse.ArgumentParser(description='SideKick365 command line interface')
# add arguments
parser.add_argument('--wordtopowerpoint', action='store_true', help='Convert a word document to a powerpoint presentation')
parser.add_argument('--wordfile', type=str, help='The word document to convert')
parser.add_argument('--powerpointfile', type=str, help='The powerpoint presentation to create')
parser.add_argument('--customphrase', type=str, help='The customization phrase to use', default="")
# parse arguments
args = parser.parse_args()
# if no arguments were specified then print help
if len(sys.argv) == 1:
    parser.print_help(sys.stderr)
    sys.exit(1)
# if the command is not wordtopowerpoint then print help
if args.wordtopowerpoint == False:
    parser.print_help(sys.stderr)
    sys.exit(1)
# if the command is wordtopowerpoint then validate the wordfile and powerpointfile were specified and are valid, and then call extractdoc(filename)
if args.wordtopowerpoint == True:
    # check the wordfile was specified
    if args.wordfile == None:
        print("Please specify a wordfile")
        sys.exit(1)
    # check the wordfile exists
    if os.path.isfile(args.wordfile) == False:
        print("The wordfile does not exist")
        sys.exit(1)
    # check the powerpointfile was specified
    if args.powerpointfile == None:
        print("Please specify a powerpointfile")
        sys.exit(1)
    # check the powerpointfile does not exist
    if os.path.isfile(args.powerpointfile) == True:
        print("The powerpointfile already exists")
        sys.exit(1)
    
    # call the function to convert the wordfile to a powerpointfile
    slidecount=sidekick365.GeneratePowerPointFromWord(args.wordfile, args.powerpointfile)
    # call the function to open the powerpoint file apply PowerPoint designer
    sidekick365.OpenPowerPointAndApplyDesigner(args.powerpointfile,slidecount)



