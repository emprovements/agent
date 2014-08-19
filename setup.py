from distutils.core import setup
import py2exe
import glob
 
setup(
	service = ["agent"],
	description = "A dummy SMTP server that logs to file.",
	cmdline_style='pywin32',
	)
