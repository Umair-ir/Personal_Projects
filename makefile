BlinkLed: BlinkLed.o
	ld -o BlinkLed BlinkLed.o
BlinkLed.o: BlinkLed.s
	as -o BlinkLed.o BlinkLed.s
