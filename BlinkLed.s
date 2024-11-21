.global _start

.equ Delay, 0x800000

.equ GPIO_BASE, 0xFE200000
.equ GPFSEL2, 0x08

.equ GPIO_OUTPUT, 0x8 

.equ GPIO_SET,  0x1C
.equ GPIO_CLEAR,0x28

.equ GPIOVAL, 0x200000 maybe @0x00200000 for gpio 21

_start:
    
    ldr r0, =GPIO_BASE

    @setting gpio 21 as output
    
    ldr r1 =GPIO_OUTPUT
    str r1, [r0, #GPFSEL2]

    #set counter
    ldr r2, =Delay

loop:
    
    @LED_ON
    ldr r1, =GPIOVAL
    str r1, [r0, #GPIO_SET] @storing in set register

    mov r10, #0   
    delay1:
        add r10, r10, #1
        cmp r10, r2
        bne, delay1

    @LED_OFF
    ldr r1, =GPIOVAL
    str r1, [r0, #GPIO_CLEAR] @storing in set register

    mov r10, #0   
    delay2:
        add r10, r10, #1
        cmp r10, r2
        bne, delay2

    b loop
