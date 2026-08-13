# -*- coding: utf-8 -*-
"""国补订单照片处理工具 - Streamlit 主程序

运行: cd deepseek && streamlit run app.py
"""
import glob
import os
import sys
from concurrent.futures import ThreadPoolExecutor, as_completed

import streamlit as st
from PIL import Image

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import image_tools as it
import xlsx_reader as xr

ROOT = os.path.dirname(os.path.abspath(__file__))
EXPORT_DIR_DEFAULT = os.path.join(os.path.expanduser("~"), "Desktop")  # 默认保存到系统桌面
UPLOAD_DIR = os.path.join(ROOT, ".upload")
PHOTO_CACHE = os.path.join(ROOT, ".photo_cache")
DEFAULT_ORDER_XLSX = os.path.join(ROOT, "新建 XLSX 工作表.xlsx")
DEFAULT_PHOTO_XLSX = (glob.glob(os.path.join(ROOT, "直供订单照片导出*.xlsx")) or [""])[0]

PARAM_DEFAULTS = {
    "rot": 0, "flip_h": False, "flip_v": False, "scale": 100,
    "bri": 1.0, "con": 1.0, "sat": 1.0, "shp": 0, "blr": 0,
    "gray": False, "inv": False, "bin": False, "thr": 128,
    "cl": 0, "cr": 0, "ct": 0, "cb": 0,
    "wm": "", "wmsz": 40, "wmop": 40, "wmpos": "右下",
    "fmt": "JPG", "q": 90,
}
POSITIONS = ["左上", "中上", "右上", "左中", "居中", "右中", "左下", "中下", "右下"]


def _init():
    st.session_state.setdefault("parse_done", False)
    st.session_state.setdefault("orders", [])
    st.session_state.setdefault("order_urls", {})
    st.session_state.setdefault("orders_no_photo", [])
    st.session_state.setdefault("orders_unmatched", [])
    st.session_state.setdefault("stats", {})
    st.session_state.setdefault("curr_index", 0)
    st.session_state.setdefault("curr_photo", 0)
    st.session_state.setdefault("marked", set())
    st.session_state.setdefault("saved", {})
    st.session_state.setdefault("sheet1_path", None)
    st.session_state.setdefault("sheet2_path", None)
    st.session_state.setdefault("uploaded_mode", False)
    st.session_state.setdefault("export_dir", EXPORT_DIR_DEFAULT)
    st.session_state.setdefault("uploaded_key", None)
    for d in (EXPORT_DIR_DEFAULT, UPLOAD_DIR, PHOTO_CACHE):
        os.makedirs(d, exist_ok=True)


def _upload_token(f):
    fid = getattr(f, "file_id", None)
    if fid:
        return fid
    return f"{f.name}|{getattr(f, 'size', 0)}"


def sidebar():
    st.sidebar.header("① 数据表选择")
    mode = st.sidebar.radio("选择方式", ["使用本地文件", "上传文件"], key="src_mode")
    sheet1_path = sheet2_path = None
    uploaded_mode = False
    if mode == "使用本地文件":
        p1 = st.sidebar.text_input("单号表路径", value=DEFAULT_ORDER_XLSX, key="p1")
        p2 = st.sidebar.text_input("直供订单照片导出表路径", value=DEFAULT_PHOTO_XLSX, key="p2")
        if p1 and p2 and os.path.exists(p1) and os.path.exists(p2):
            sheet1_path, sheet2_path = p1, p2
        else:
            st.sidebar.warning("路径无效，请检查文件是否存在")
    else:
        f1 = st.sidebar.file_uploader("上传 单号表 (.xlsx)", type=["xlsx"], key="f1")
        f2 = st.sidebar.file_uploader("上传 直供订单照片导出表 (.xlsx)", type=["xlsx"], key="f2")
        if f1 and f2:
            uploaded_mode = True
            sheet1_path = os.path.join(UPLOAD_DIR, os.path.basename(f1.name) or "order.xlsx")
            sheet2_path = os.path.join(UPLOAD_DIR, os.path.basename(f2.name) or "photo.xlsx")
            token = f"{_upload_token(f1)}|{_upload_token(f2)}"
            if st.session_state.uploaded_key != token:
                os.makedirs(UPLOAD_DIR, exist_ok=True)
                with open(sheet1_path, "wb") as fp:
                    fp.write(f1.getbuffer())
                with open(sheet2_path, "wb") as fp:
                    fp.write(f2.getbuffer())
                st.session_state.uploaded_key = token
    st.sidebar.header("② 输出设置")
    export_dir = st.sidebar.text_input("图片保存目录", value=st.session_state.export_dir, key="export_dir_input")
    st.session_state.export_dir = (export_dir or EXPORT_DIR_DEFAULT).strip() or EXPORT_DIR_DEFAULT
    st.sidebar.caption("保存命名：订单编号_照片序号.扩展名")
    st.sidebar.button("解析并匹配数据", type="primary", width="stretch", key="btn_parse")
    return sheet1_path, sheet2_path, uploaded_mode


def do_parse(sheet1_path, sheet2_path):
    if not sheet1_path or not sheet2_path or not os.path.exists(sheet1_path) or not os.path.exists(sheet2_path):
        st.error("请先正确选择两个数据表文件")
        return False
    try:
        with st.spinner("正在解析单号表（订单编号）..."):
            order_list = xr.read_order_numbers(sheet1_path)
        with st.spinner("正在流式解析直供订单照片导出表（约 1 分钟）..."):
            result = xr.parse_direct_table(sheet2_path, set(order_list))
    except Exception as e:
        st.error(f"解析失败：{e}")
        return False
    photo_orders = result["photo_orders"]
    seen = result["seen_orders"]
    with_photo = [o for o in order_list if o in photo_orders]
    no_photo = [o for o in order_list if o in (seen - set(photo_orders))]
    unmatched = [o for o in order_list if o not in seen]
    st.session_state.update(
        parse_done=True, sheet1_path=sheet1_path, sheet2_path=sheet2_path,
        orders=with_photo, order_urls=photo_orders,
        orders_no_photo=no_photo, orders_unmatched=unmatched,
        stats={"total": len(order_list), "with_photo": len(with_photo),
               "no_photo": len(no_photo), "unmatched": len(unmatched),
               "matched_rows": result["matched_rows"], "total_rows": result["total_rows"],
               "len_counter": result["len_counter"]},
        curr_index=0, curr_photo=0, marked=set(), saved={},
    )


def get_order_photos(order):
    """按订单并发下载全部照片到本地缓存（含进度条），结果存入 session_state。"""
    key = ("photos", order)
    if key not in st.session_state:
        urls = st.session_state.order_urls[order]
        prog = st.progress(0.0, text=f"正在下载订单 {order} 的照片 ...")
        results = {}
        with ThreadPoolExecutor(max_workers=6) as ex:
            futs = {ex.submit(it.ensure_downloaded, u, PHOTO_CACHE): u for u in urls}
            done = 0
            for fut in as_completed(futs):
                results[futs[fut]] = fut.result()
                done += 1
                prog.progress(done / len(urls), text=f"正在下载订单 {order} 的照片 ... {done}/{len(urls)}")
        st.session_state[key] = results
    return st.session_state[key]


def get_original(order, ph):
    """加载指定订单第 ph 张原图（缓存于 session_state）。下载失败返回 None。"""
    key = ("img", order, ph)
    if key not in st.session_state:
        u = st.session_state.order_urls[order][ph]
        path, ok = get_order_photos(order).get(u, (None, False))
        st.session_state[key] = it.load_image(path) if ok else None
    return st.session_state[key]


def advance_after_save(n_urls, n_orders, idx):
    """保存后自动跳到下一张/下一单。"""
    if st.session_state.curr_photo < n_urls - 1:
        st.session_state.curr_photo += 1
    elif idx < n_orders - 1:
        st.session_state.curr_index = idx + 1
        st.session_state.curr_photo = 0
    else:
        st.session_state.curr_index = 0
        st.session_state.curr_photo = 0


def main():
    _init()
    sheet1_path, sheet2_path, uploaded_mode = sidebar()

    st.title("🖼️ 国补订单照片处理工具")
    st.caption("流程：选择两张数据表 → 解析匹配 → 逐单处理图片 → 保存 / 标记已上传")

    if st.session_state.btn_parse:
        if do_parse(sheet1_path, sheet2_path):
            st.session_state.uploaded_mode = uploaded_mode
            st.rerun()

    if not st.session_state.parse_done:
        st.info("👈 请在左侧选择单号表与直供订单照片导出表，然后点击“解析并匹配数据”。")
        return

    stats = st.session_state.stats
    m = st.columns(5)
    m[0].metric("单号表订单总数", stats.get("total", 0))
    m[1].metric("✅ 匹配且有照片", stats.get("with_photo", 0))
    m[2].metric("匹配但无照片", stats.get("no_photo", 0))
    m[3].metric("未匹配订单", stats.get("unmatched", 0))
    m[4].metric("直供表数据行", stats.get("total_rows", 0))
    st.caption(f"单号表：{st.session_state.sheet1_path}　|　图片输出：{st.session_state.export_dir}")

    with st.expander(f"未匹配订单（{len(st.session_state.orders_unmatched)} 个，不在直供表三方单号中）"):
        st.write("、".join(map(str, st.session_state.orders_unmatched)) or "无")
    with st.expander(f"匹配但无照片的订单（{len(st.session_state.orders_no_photo)} 个）"):
        st.write("、".join(map(str, st.session_state.orders_no_photo)) or "无")

    orders = st.session_state.orders
    if not orders:
        st.warning("没有找到任何匹配且有照片的订单，请检查两张数据表是否选择正确。")
        return
    idx = min(st.session_state.curr_index, len(orders) - 1)
    st.session_state.curr_index = idx
    order = orders[idx]
    urls = st.session_state.order_urls[order]

    st.divider()
    hc = st.columns([4, 1, 1])
    hc[0].markdown(f"### 📋 订单 {idx + 1} / {len(orders)}")
    if hc[1].button("‹ 上一单", key="prev_order", width="stretch"):
        st.session_state.curr_index = max(0, idx - 1)
        st.session_state.curr_photo = 0
        st.rerun()
    if hc[2].button("下一单 ›", key="next_order", width="stretch"):
        st.session_state.curr_index = min(len(orders) - 1, idx + 1)
        st.session_state.curr_photo = 0
        st.rerun()

    _urls = st.session_state.order_urls
    sel = st.selectbox("跳转到订单", range(len(orders)), index=idx,
                       format_func=lambda i, _orders=orders, _urls=_urls: f"{_orders[i]}（{len(_urls[_orders[i]])}张照片）",
                       key="order_jump")
    if sel != idx:
        st.session_state.curr_index = sel
        st.session_state.curr_photo = 0
        st.rerun()



    ph = min(st.session_state.curr_photo, len(urls) - 1)
    st.session_state.curr_photo = ph
    st.markdown(f"#### 🖼️ 该订单共 {len(urls)} 张照片")

    thumbs = get_order_photos(order)
    tcols = st.columns(min(len(urls), 6))
    for i, u in enumerate(urls):
        path, ok = thumbs.get(u, (None, False))
        with tcols[i % len(tcols)]:
            if ok:
                cap = f"第{i+1}张" + (" ✅已存" if (order, i) in st.session_state.saved else "")
                st.image(Image.open(path).convert("RGB"), caption=cap, width=150)
            else:
                st.warning(f"第{i+1}张下载失败")

    _saved = st.session_state.saved
    psel = st.radio("选择要处理的照片", range(len(urls)), index=ph,
                    format_func=lambda i, _order=order, _saved=_saved: f"第 {i+1} 张" + ("（已保存）" if (_order, i) in _saved else ""),
                    horizontal=True, key="photo_sel")
    if psel != ph:
        st.session_state.curr_photo = psel
        st.rerun()

    orig = get_original(order, ph)
    if orig is None:
        st.error(f"第 {ph+1} 张照片下载失败，无法处理。")
        if st.button("重新下载本单照片"):
            st.session_state.pop(("photos", order), None)
            st.session_state.pop(("img", order, ph), None)
            st.rerun()
        return

    pf = f"{order}|{ph}"
    wc, pc = st.columns([1, 1], gap="large")
    with wc:
        st.subheader("🛠️ 图片处理")
        with st.expander("旋转 / 翻转 / 缩放", expanded=True):
            st.slider("旋转角度 (°)", -180, 180, 0, key=f"rot_{pf}")
            st.checkbox("水平翻转", key=f"flip_h_{pf}")
            st.checkbox("垂直翻转", key=f"flip_v_{pf}")
            st.slider("缩放比例 (%)", 10, 300, 100, key=f"scale_{pf}")
        with st.expander("裁剪（按百分比）", expanded=False):
            st.slider("裁剪左侧 (%)", 0, 90, 0, key=f"cl_{pf}")
            st.slider("裁剪右侧 (%)", 0, 90, 0, key=f"cr_{pf}")
            st.slider("裁剪上侧 (%)", 0, 90, 0, key=f"ct_{pf}")
            st.slider("裁剪下侧 (%)", 0, 90, 0, key=f"cb_{pf}")
        with st.expander("亮度 / 对比度 / 饱和度 / 锐化", expanded=False):
            st.slider("亮度", 0.2, 2.5, 1.0, 0.05, key=f"bri_{pf}")
            st.slider("对比度", 0.2, 2.5, 1.0, 0.05, key=f"con_{pf}")
            st.slider("饱和度", 0.0, 2.5, 1.0, 0.05, key=f"sat_{pf}")
            st.slider("锐化强度", 0, 100, 0, key=f"shp_{pf}")
        with st.expander("滤镜 / 二值化", expanded=False):
            st.slider("高斯模糊半径", 0, 20, 0, key=f"blr_{pf}")
            st.checkbox("灰度", key=f"gray_{pf}")
            st.checkbox("反色", key=f"inv_{pf}")
            bin_on = st.checkbox("二值化", key=f"bin_{pf}")
            st.slider("二值化阈值", 0, 255, 128, disabled=not bin_on, key=f"thr_{pf}")
        with st.expander("文字水印", expanded=False):
            st.text_input("水印文字（留空则不添加）", key=f"wm_{pf}")
            st.slider("字号", 10, 200, 40, key=f"wmsz_{pf}")
            st.slider("透明度 (%)", 5, 100, 40, key=f"wmop_{pf}")
            st.selectbox("位置", POSITIONS, index=POSITIONS.index("右下"), key=f"wmpos_{pf}")
        with st.expander("输出格式", expanded=False):
            st.selectbox("格式", ["JPG", "PNG", "WEBP"], key=f"fmt_{pf}")
            st.slider("JPG/WEBP 质量", 30, 100, 90, key=f"q_{pf}")

    params = {k: st.session_state.get(f"{k}_{pf}", d) for k, d in PARAM_DEFAULTS.items()}
    processed = it.apply_transforms(orig, params)

    with pc:
        st.subheader("👁️ 预览")
        c1, c2 = st.columns(2)
        c1.caption("原图")
        c1.image(orig, width="stretch")
        c2.caption("处理后")
        c2.image(processed, width="stretch")

    st.divider()
    done = (order, ph) in st.session_state.saved
    bc = st.columns([2, 1, 1, 1])
    auto_next = bc[0].checkbox("保存后自动跳到下一张/下一单", value=True, key="auto_next")
    if bc[1].button("💾 保存当前图片" + ("（已保存·可覆盖）" if done else ""), type="primary", width="stretch"):
        fmt_ = params["fmt"]
        ext = it.FORMAT_EXT.get(fmt_, "jpg")
        outdir = st.session_state.export_dir
        os.makedirs(outdir, exist_ok=True)
        name = f"{order}_{ph + 1}.{ext}"
        path = os.path.join(outdir, name)
        it.save_image(processed, path, fmt_, params["q"])
        st.session_state.saved[(order, ph)] = path
        st.toast(f"✅ 已保存：{name}")
        if auto_next:
            advance_after_save(len(urls), len(orders), idx)
            st.rerun()
    if bc[2].button("↩️ 重置本图参数", width="stretch"):
        for k, v in PARAM_DEFAULTS.items():
            st.session_state[f"{k}_{pf}"] = v
        st.rerun()
    if bc[3].button("⬇️ 下载此图", width="stretch"):
        fmt_ = params["fmt"]
        st.download_button("点击下载", data=it.image_to_bytes(processed, fmt_, params["q"]),
                           file_name=f"{order}_{ph + 1}.{it.FORMAT_EXT.get(fmt_, 'jpg')}",
                           key=f"dl_{pf}")

    st.divider()
    mrow = st.columns([2, 3])
    marked = order in st.session_state.marked
    mrow[0].markdown("**上传状态：** " + ("✅ 已标记为已上传" if marked else "⬜ 未标记"))
    if mrow[1].button("📤 标记为已上传（单号表·是否上传=是）", width="stretch", disabled=marked):
        path1 = st.session_state.sheet1_path
        try:
            with st.spinner("正在写入单号表..."):
                n = xr.mark_uploaded(path1, order)
            if n:
                st.session_state.marked.add(order)
                st.success(f"已将订单 {order} 的“是否上传”列写入“是”（共更新 {n} 行）")
                if st.session_state.uploaded_mode:
                    with open(path1, "rb") as fp:
                        st.download_button("⬇️ 下载已更新的单号表", data=fp.read(),
                                           file_name=os.path.basename(path1),
                                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            else:
                st.warning(f"在单号表中未找到订单编号 {order}，未写入")
        except Exception as e:
            st.error(f"写入失败：{e}")


if __name__ == "__main__":
    main()

