#!/usr/bin/python
# -*- coding: utf-8 -*-

#    Copyright 2014, Dennis Drescher
#    All rights reserved.
#
#    This library is free software; you can redistribute it and/or modify
#    it under the terms of the GNU Lesser General Public License as published
#    by the Free Software Foundation; either version 2.1 of License, or
#    (at your option) any later version.
#
#    This program is distributed in the hope that it will be useful,
#    but WITHOUT ANY WARRANTY; without even the implied warranty of
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
#    Lesser General Public License for more details.
#
#    You should also have received a copy of the GNU Lesser General Public
#    License along with this library in the file named "LICENSE".
#    If not, write to the Free Software Foundation, 51 Franklin Street,
#    suite 500, Boston, MA 02110-1335, USA or visit their web page on the
#    internet at http://www.fsf.org/licenses/lgpl.html.

# Description
#
# This module provides a simple Python interface to various ImageMagick
# functions it will access through the subprocess module.


# Import modules
import sys, os, shutil, subprocess, tempfile


class ImgProcess (object) :

	def __init__ (self, parent=None) :
		self.tools          = tools.Tools()


	###############################################################################
	############################## General Functions ##############################
	###############################################################################

	def outlinePic (self, inFile) :
		'''Add a simple outline to a picture. Return the name of the file
		that was just created.'''

		(name, ext)     = os.path.splitext(inFile)
		outFile         = name + '-border.jpg'
		cmd = ['convert', inFile, '-bordercolor', 'black', '-border', '1x1', outFile]

		# Run the command
		try :
			rCode = subprocess.call(cmd)
			return outFile
		except Exception as e :
			self.tools.sendError('Imagemagick outline failed with: ' + str(e))


	def processPicFile (self, inFile, outFile, rotate, size, caption, outline=False, viewer=False, compress=True) :
		'''Prepare the arguments for an Imagemagick process and then run the process.'''

		# Make tempfile
		workFile        = tempfile.NamedTemporaryFile().name + '.jpg'
		shutil.copyfile(inFile, workFile)

		# Add an outline to a pic
		if self.tools.str2bool(outline) :
			workFile        = self.outlinePic(workFile)

		# Begin the output command set
		cmds = ['convert']
		# Set the output size
		sizeDim = '400x300'
		fontSize = 18
		if size :
			if size.lower() == 'small' :
				sizeDim = '400x300'
				fontSize = 18
			elif size.lower() == 'medium' :
				sizeDim = '800x600'
				fontSize = 24
			elif size.lower() == 'large' :
				sizeDim = '1024x768'
				fontSize = 28
		# Need to append the caption now if there is one
		if caption :
			cmds.append('-caption')
			cmds.append(caption)
		# Now tack on the input file
		cmds.append(workFile)
		# Build the rest of the command set
		base = ['-thumbnail', sizeDim, '-font', 'Andika-Basic-Regular', \
			'-pointsize', str(fontSize), '-border', '2x2', '-density', '72', \
				'-gravity', 'center', '-bordercolor', 'white', '-background', 'black', \
					'-polaroid', str(rotate), outFile]
		for c in base :
			cmds.append(c)

		# Run the command
		try :
			rCode = subprocess.call(cmds)
		except Exception as e :
			self.sendError('Imagemagick failed with: ' + str(e))

		# View the results (os.system allows the terminal to return right away)
		if self.tools.str2bool(viewer) :
			os.system('eog ' + outFile + ' &')
