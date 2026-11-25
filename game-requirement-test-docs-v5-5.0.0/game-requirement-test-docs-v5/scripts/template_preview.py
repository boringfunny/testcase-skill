"""Auxiliary diagram of exported categories. Not a native XMind rendering."""
from PIL import Image, ImageDraw
import math


def overview(root, path, font):
    branches=root.get('children',{}).get('attached',[])
    half=math.ceil(len(branches)/2)
    height=max(600,160+max(half,len(branches)-half)*130)
    im=Image.new('RGB',(1600,height),'#eee8e6');d=ImageDraw.Draw(im)
    d.text((24,20),'XMind 分类辅助示意（非原生渲染；下级分组默认折叠）',font=font(18),fill='#88675e')
    def box(x,y,w,text,primary=False):
        d.rounded_rectangle((x,y-30,x+w,y+30),12,fill='#c1aba5' if primary else '#e6ddda',outline='#88675e',width=2)
        size=22
        while d.textlength(text,font=font(size))>w-24 and size>11:size-=1
        d.text((x+12,y-size//2-3),text,font=font(size),fill='#2d221f')
    box(620,height//2,360,root['title'],True)
    def count(t):
        return (1 if t.get('labels') else 0)+sum(count(c) for c in t.get('children',{}).get('attached',[]))
    for i,b in enumerate(branches):
        right=i<half;index=i if right else i-half
        y=130+index*130;x=1120 if right else 100
        start=(980 if right else 620,height//2);end=(x if right else x+380,y)
        d.line([start,((start[0]+end[0])//2,start[1]),((start[0]+end[0])//2,y),end],fill='#88675e',width=2)
        box(x,y,380,b['title'],True)
        d.text((x+12,y+38),f'{count(b)} 条测试点 · 展开后查看简述',font=font(17),fill='#2d221f')
    im.save(path)
