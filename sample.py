# -*- coding: utf-8 -*-

# Sample of game macro.
# It will only work on my PC.

import time
from macrommon import *

pahwnd, chhwnd = getHwnd('LDPlayerMainFrame', 'LDPlayer', 'RenderWindow', 'TheRender')

tol   = 5
inter = 1

def swap(run):
    randClick(chhwnd, 722, 336) # Form

    colorWait(chhwnd, 107, 441, (210, 224,  92), inter=inter)    # E1.M1.Life
    randClick(chhwnd, 262, 265, inter=inter) # E1.M2
    colorWait(chhwnd, 656,  79, (249, 176,   0))    # R1C6.Star

    if run % 2 == 0:
        randClick(chhwnd, 286, 532, inter=inter)    # R3C3 @UZI
    else:
        randClick(chhwnd, 188, 160, inter=inter)    # R1C2 @HK416

    colorWait(chhwnd, 107, 441, (210, 224,  92), inter=inter)    # E1.M1.Life

    randClick(chhwnd, 135,  33) # TELEPORT
    randClick(chhwnd,  89, 145) # TELEPORT.Battle

def ep(repeat=1):
    for x in range(repeat):

        #
        # Open
        #

        colorWait(chhwnd, 253,  72, (230,  53,   0))    # CM.DIFFICULTY.EMERGENCY
        randClick(chhwnd, 528, 404)
        colorWait(chhwnd, 419, 415, (253, 179,   0))
        randClick(chhwnd, 467, 436)
        colorWait(chhwnd,  22, 590, (240,   0,  77))
        
        time.sleep(inter << 2)

        sendKey(chhwnd, 0x73, inter=inter)      # VK_F4
        randDrag(chhwnd, 100, 300, 700, 300)    # Align

        randClick(chhwnd, 434, 421, tol=tol) # 07:00 PORT
        colorWait(chhwnd, 219, 416, (255, 207,   0))    # E1.M2.Ammo

        if colorMatch(chhwnd, ((560 + 649) // 2), 399, (  0,   0,   0)):
            randClick(chhwnd, 605, 305)
            randClick(chhwnd, 576, 390)

        randClick(chhwnd, 726, 472, inter=inter)
        colorWait(chhwnd,  22, 590, (240,   0,  77))
        randClick(chhwnd, 404, 212, tol=tol) # HQ
        colorWait(chhwnd, 106, 399, (210, 224,  92))    # E2.M1.Life
        randClick(chhwnd, 726, 472, inter=inter)
        colorWait(chhwnd,  22, 590, (240,   0,  77))
        
        randClick(chhwnd, 709, 560)

        time.sleep(inter << 2)

        #
        # Middle
        #

        randClick(chhwnd, 404, 212, tol=tol, inter=inter) # HQ @E2
        randClick(chhwnd, 404, 212, tol=tol, inter=inter) # HQ @E2
        colorWait(chhwnd, 106, 399, (210, 224,  92))    # E2.M1.Life
        randClick(chhwnd, 737, 425, inter=inter)
        colorWait(chhwnd,  22, 590, (240,   0,  77))
        
        for _ in range(6):
            randClick(chhwnd,  22, 590) # Deselect

        randClick(chhwnd,  12, 519)
        randClick(chhwnd, 434, 421, tol=tol) # 07:00 PORT @E1
        randClick(chhwnd, 492, 302, tol=tol) # ILOC 1
        randClick(chhwnd, 550, 302, tol=tol) # ILOC 2
        randClick(chhwnd, 740, 560)

        colorWait(chhwnd, 103,  67, ( 90, 204, 219), inter=inter << 2)

        #
        # Close
        #

        colorWait(chhwnd, 103,  67, ( 90, 204, 219))

        swap(x)

        print('Done {:2d}'.format(x + 1))

if __name__ == '__main__':
    repeat = int(input('REPEAT : '))

    print('START')

    ep(repeat)
    time.sleep(3)

    print('END')

    randClick(chhwnd,  62,  32) # Back
