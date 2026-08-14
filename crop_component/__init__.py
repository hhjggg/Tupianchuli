# -*- coding: utf-8 -*-
"""微信截图式图片裁剪 + 点击放大自定义组件。

用法:
    from crop_component import image_crop
    res = image_crop(img=base64_dataurl, width=460, reset=0, key="crop_xxx")
    # res = {zoom, crop:{x,y,w,h}(原始像素), confirmed, canceled}
"""
import os

import streamlit.components.v1 as components

_FRONTEND_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "frontend")

image_crop = components.declare_component(
    "wechat_image_crop",
    path=_FRONTEND_DIR,
)
