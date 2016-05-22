if len(sys.argv) == 3:
	node = sys.argv[0]
	server = sys.argv[1]
	action = sys.argv[2]

	print action," Application Server ", server, " on node ", node
	if action == 'start':
		AdminControl.startServer(server, node)
	if action == 'stop':
		AdminControl.stopServer(server, node)
else:
	print "Syntax: <scriptname> <server> {start|stop}"