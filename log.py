import logging
import colorlog

global logger

def init():
	global logger

	try:
		# logHandler = TimedRotatingFileHandler("home_automation.log",when="D", interval=1)
		logHandler = colorlog.StreamHandler()
		formatter = colorlog.ColoredFormatter(
			"%(asctime)s - %(log_color)s%(levelname)-8s - %(white)s%(message)s",
			datefmt=' %d-%m-%Y %H:%M:%S',
			reset=True,
			log_colors={
			    'DEBUG':    'cyan',
			    'INFO':     'green',
			    'WARNING':  'yellow',
			    'ERROR':    'red',
			    'CRITICAL': 'red',
			}
		)

		logHandler.setFormatter( formatter )
		logger = colorlog.getLogger('example')
		logger.addHandler( logHandler )
		logger.setLevel( logging.DEBUG )
	except: 
		print("Error initializing logging...Cannot continue")
		raise SystemExit

def debug(message):
    logger.debug(message)

def info(message):
    logger.info(message)

def warning(message):
    logger.warning(message)

def error(message):
    logger.error(message)

def critical(message):
    logger.critical(message)

