-- title:   Mini-Game
-- author:  Robert Kuropkat
-- desc:    Semester 01 project
-- site:    https://doersofstuff.com/
-- license: MIT License (change this to your license of choice)
-- version: 0.1
-- script:  lua

PALETTE_MAP=0x3FF0

x   = 60  -- starting x-position
y   = 60  -- starting y-position
spd = 1   -- speed
b   = 1   -- beep toggle
f   = 0   -- flash toggle
s   = 1   -- sprite scale

function TIC()

	cls(0)

 b=b+0.2  -- set beep speed
 f=f+0.1  -- set flash speed

 x=math.max(0,math.min(232,x))
 y=math.max(0,math.min(128,y))

 if     btn(0) then -- up
               poke4(PALETTE_MAP*2+10,3)
               sfx(math.floor(b)%2)
               y=y-spd 
               s=math.floor(f)%2
 elseif btn(1) then -- down
               poke4(PALETTE_MAP*2+10,1)
               sfx(math.floor(b)%2)
               y=y+spd 
               s=math.floor(f)%2
 elseif btn(2) then -- left
               poke4(PALETTE_MAP*2+10,0)
               sfx(math.floor(b)%2)
               x=x-spd 
               s=math.floor(f)%2
 elseif btn(3) then -- right
               poke4(PALETTE_MAP*2+10,6)
               sfx(math.floor(b)%2)
               x=x+spd 
               s=math.floor(f)%2
 else          poke4(PALETTE_MAP*2+10,10) 
               sfx(-1)
               s=1
 end

	spr( 1  -- index
	   , x  -- x-position
				, y  -- y-position
				, 0  -- color key
				, s  -- scale
				, 0		-- flip
				, 0  -- rotate
				, 2  -- width in sprites
				, 2  -- height in sprites
				)

end

-- <TILES>
-- 001:eccccccccc888888caaaaaaaca888888cacccccccacc0ccccacc0ccccacc0ccc
-- 002:ccccceee8888cceeaaaa0cee888a0ceeccca0ccc0cca0c0c0cca0c0c0cca0c0c
-- 003:eccccccccc888888caaaaaaaca888888cacccccccacccccccacc0ccccacc0ccc
-- 004:ccccceee8888cceeaaaa0cee888a0ceeccca0cccccca0c0c0cca0c0c0cca0c0c
-- 017:cacccccccaaaaaaacaaacaaacaaaaccccaaaaaaac8888888cc000cccecccccec
-- 018:ccca00ccaaaa0ccecaaa0ceeaaaa0ceeaaaa0cee8888ccee000cceeecccceeee
-- 019:cacccccccaaaaaaacaaacaaacaaaaccccaaaaaaac8888888cc000cccecccccec
-- 020:ccca00ccaaaa0ccecaaa0ceeaaaa0ceeaaaa0cee8888ccee000cceeecccceeee
-- </TILES>

-- <WAVES>
-- 000:00000000ffffffff00000000ffffffff
-- 001:0123456789abcdeffedcba9876543210
-- 002:0123456789abcdef0123456789abcdef
-- </WAVES>

-- <SFX>
-- 000:00a0f000f000f000f000f000f000f000f000f000f000f000f000f000f000f000f000f000f000f000f000f000f000f000f000f000f000f000f000f000100000000000
-- 001:000000000000000000000000000000000000000000000000000000900090009000000000000000000000000000000000000000000000000000000000400000000000
-- </SFX>

-- <PALETTE>
-- 000:1a1c2c5d275db13e53ef7d57ffcd75a7f07038b76425717929366f3b5dc941a6f673eff7f4f4f494b0c2566c86333c57
-- </PALETTE>

