# pypiprocessbook
Tool for request data from PI Processbook

Usage:
    from pypiprocessbook import PI

    pi = PI("server_name")

    results = pi.read("tag", start_date, end_date, frequency)