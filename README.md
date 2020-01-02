# pypiprocessbook
Tool for request data from PI Processbook

Usage:

    from pypiprocessbook import PI

    pi = PI("server_name")

    results = pi.read("tag") # get the last value for the tag
    results = pi.read("tag", start_time) # get the value for the tag at specified datetime
    results = pi.read("tag", start_time, end_time, frequency) # get the values for the tag at specified datetime interval every 'frequency' minutes