"""
This is the Call Process Script.
Used to make managing calling logs all the simpler.

The Directory.py file holds a dictionary of frequently called extensions. This .py file
is used to provide the data for the option of automatically inserting extension owners.
This file can be modified at any time to reflect extension changes. The script will also
query the user during runtime if they would like to utilize the automatic insertion feature.

This script uses DataNitro (a Python plug-in for Excel), the UltraEdit text editor,
the collapse_columns1.js file, and the 'VBA Macros.txt' file.

Instructions on set-up and use are included in the 'README.txt' file.
Documentation of methods and macros can be found in the 'DOCUMENTATION.txt' file.

Created by Kevin Pietrow, circa 2012-2013
"""

'''
Error Checking Needed:
Test with different data sizes
Test with data from different months
'''

import Directory

def extension_cleanup():
	"""Total and remove x2000s, then add total to x3000"""
	count = 2
	total = 0
	
	while 1:
		extension = Cell("A" + str(count)).value
		
		if extension is None:
			Cell("B" + str(count - 1)).set_active()
			break
		
		else:
			if extension >= 2000 and extension < 3000:	# Search x2000 extensions
				total += Cell("B" + str(count)).value	# Total, then remove
				Cell("B" + str(count)).clear()
				Cell("A" + str(count)).clear()
		
			elif extension == 3000:	# Search for operator extension
				Cell("B" + str(count)).value += total
		
			if count % 100 == 0:	# Every 100 records, print status report
				print str(count) + " records have been processed..."
			
		count += 1
	
	return
	
def extension_directory():
	"""Uses directory to fill in extension owners"""
	count = 2
	directory = Directory.return_directory()
	# KeyError
	while count < 12:
		try:
			var = directory[str(Cell("A" + str(count)).value)]
		except KeyError:
			var = "Unavailable"
		finally:
			Cell("B" + str(count)).value = var
			count += 1
	return
	
def call_recipients():
	"""Query user if they want to automatically put in extension owners"""
	answer = raw_input("Would you like to automatically enter the extension owners? (y/n):  ").lower()
	while 1:
		if answer == "y":
			# Label top ten callers, and graph results
			extension_directory()
			break

		elif answer == "n":
			# Wait until user has entered data
			raw_input("Press <enter> when you have finished entering the information")
			break
		
		else:
			answer = raw_input("Not a valid response, please enter again: ").lower()

	print "You can update the automatic directoy in the DIRECTORY.py file"
	return



	
def main():
	print "Call Processer is beginning"

	# construct first pivot table, set to active sheet
	VBA("Date_pivot")
	print "First pivot table complete.\n"

	active_sheet("Sheet3")

	# Total, then remove x2000 extensions
	extension_cleanup()
	print "x2000 extensions have been totaled and erased"

	# Sort remaining records
	VBA("Largest_to_Smallest")
	print "\nOrganized records and inserted column\n"
	
	call_recipients()
	
	VBA("First_Graph")
	print "\nSuccessfully created first graph"

	# Make Daily Calls graph
	VBA("Daily_Calls")
	active_sheet("Sheet5")
	print "Daily Calls graph completed."

	VBA("Average_Hourly")
	print "Average Hourly graph complete"

	VBA("Cleanup")
	print "Formatted and labeled sheets."
	
	print "\nSuccessfuly completed all tasks. Thank you for using the Call Processor for all of your call processing needs. Please disable the DataNitro Excel add-in until you need use of the Call Processer again."
	raw_input("Press <enter> to exit")
	
	return

if __name__ == '__main__':
	main()