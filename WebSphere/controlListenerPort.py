if len(sys.argv) == 3:
	node = sys.argv[0]
	process = sys.argv[1]
	action = sys.argv[2]

	print action," Listener Port of ", process

	lPorts = AdminControl.queryNames('type=ListenerPort,node='+node+',process='+process+',*')
	# print lPorts

	# get line separator 
	import java
	lineSeparator = java.lang.System.getProperty('line.separator')

	lPortsArray = lPorts.split(lineSeparator)
	for lPort in lPortsArray:
		state = AdminControl.getAttribute(lPort, 'started')
		if state == 'false' and action == 'start':
			AdminControl.invoke(lPort, 'start')
			print "Listener Port started"
		if state == 'true' and action == 'stop':
			AdminControl.invoke(lPort, 'stop')
			print "Listener Port stopped"
else:
	print "Syntax: <scriptname> <node> <process> {start|stop}"