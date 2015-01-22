import sys
import re
import os
import mechanize
from bs4 import BeautifulSoup

GRADER_BASE = "http://grader.eng.src.ku.ac.th"
SOURCE_HEADER = "Content-Disposition"
HEADER_MATCH = "attachment.*filename=\\\"?([^\"]*)\\\"?"

def run(args):
	if len(args) < 2:
		print("usage: %s <user file> [output directory]" % (args[0]))
		print("User file is a list of user information in this format: user,pwd[,output_name]")
		print("Any invalid line will be skipped")
		return
	parent_dir = "output"
	if len(args) > 1:
		parent_dir = args[2]

	f = open(args[1], "r")
	if f is None:
		print("Error occured while reading a file")
		return
	user_list = f.read().split("\n")
	f.close()

	for user in user_list:
		user_info = user.split(",")
		if len(user_info) < 2:
			continue
		username = user_info[0]
		password = user_info[1]
		outputname = username
		if len(user_info) > 2:
			outputname = user_info[2]

		br = mechanize.Browser()
		br.set_handle_robots(False)
		try:
			br.open(GRADER_BASE)
		except Exception as msg:
			print("Grader Error! %s" % (msg))
			return
		if len(list(br.forms())) < 1:
			print("No login form in grader... Please check the grader...")
			return
		br.form = list(br.forms())[0]
		br.form["login"] = username
		br.form["password"] = password
		print("Logging into grader as \"%s\"..." % (username))
		br.submit()

		index_html = BeautifulSoup(br.response().read())

		download_links = set()
		for link in index_html.find_all(name="a", text="[src]"):
			download_links.add(link.get("href"))

		if len(download_links) <= 0:
			print("[%s] No download link" % (username))
			continue

		for link in download_links:
			print("[%s] Downloading %s..." % (username, link))
			try:
				response = br.open(GRADER_BASE+link)
			except Exception as e:
				print("[%s] Grader Error! %s" % (username, e))
				continue
			headers = response.info()
			if SOURCE_HEADER not in headers:
				print("[%s] No desired header. Session timeout or invalid link maybe?" % (username))
				continue
			header_match = re.search(HEADER_MATCH, headers[SOURCE_HEADER])
			if header_match is None:
				print("[%s] Invalid header" % (username))
				continue
			if header_match.groups < 1:
				filename = os.path.basename(link)
				print("[%s] Warning! Potential invalid header match pattern" % (username))
			else:
				filename = header_match.group(1)
			full_path = [parent_dir, outputname]
			try:
				os.makedirs(os.path.join(*full_path))
			except Exception:
				pass
			full_path.append(filename)
			print("[%s] Saving code as \"%s\" to \"%s\"" % (username, filename, outputname))
			f = open(os.path.join(*full_path), "w")
			f.write(response.read())
			f.close()



if __name__ == "__main__":
	run(sys.argv)
