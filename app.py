# -*- coding: utf-8 -*-
"""国补订单照片处理工具 - Streamlit 主程序

运行: cd deepseek && streamlit run app.py
"""
import base64
import glob
import io
import os
import sys
from concurrent.futures import ThreadPoolExecutor, as_completed

import streamlit as st
from PIL import Image

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from crop_component import image_crop
import image_tools as it
import xlsx_reader as xr

ROOT = os.path.dirname(os.path.abspath(__file__))
EXPORT_DIR_DEFAULT = os.path.join(os.path.expanduser("~"), "Desktop")  # 默认保存到系统桌面
UPLOAD_DIR = os.path.join(ROOT, ".upload")
PHOTO_CACHE = os.path.join(ROOT, ".photo_cache")
DEFAULT_ORDER_XLSX = os.path.join(ROOT, "新建 XLSX 工作表.xlsx")
DEFAULT_PHOTO_XLSX = (glob.glob(os.path.join(ROOT, "直供订单照片导出*.xlsx")) or [""])[0]


def _img_dataurl(img, fmt="JPEG"):
    """PIL 图片 → base64 data URL（传给裁剪组件）。"""
    buf = io.BytesIO()
    img.save(buf, fmt)
    return "data:image/jpeg;base64," + base64.b64encode(buf.getvalue()).decode("utf-8")


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
    st.session_state.setdefault("uploaded_paths", {})
    st.session_state.setdefault("merged_path", None)
    st.session_state.setdefault("crop_box", {})
    st.session_state.setdefault("crop_reset", {})
    st.session_state.setdefault("editor_open", None)
    for d in (EXPORT_DIR_DEFAULT, UPLOAD_DIR, PHOTO_CACHE):
        os.makedirs(d, exist_ok=True)


def _upload_token(f):
    fid = getattr(f, "file_id", None)
    if fid:
        return fid
    return f"{f.name}|{getattr(f, 'size', 0)}"


def _save_upload(f, prefix):
    """把上传文件落盘到 .upload/ 目录，返回本地路径。"""
    name = os.path.basename(f.name) or f"{prefix}.xlsx"
    path = os.path.join(UPLOAD_DIR, f"{prefix}_{name}")
    with open(path, "wb") as fp:
        fp.write(f.getbuffer())
    return path


def sidebar():
    st.sidebar.header("① 数据表选择")
    mode = st.sidebar.radio("选择方式", ["使用本地文件", "上传文件"], key="src_mode")
    sheet1_path = None
    done_files = []   # 处理完成的直供订单表（1 个）
    raw_files = []    # 未处理的直供订单照片导出表（多个）
    uploaded_mode = False
    if mode == "使用本地文件":
        p1 = st.sidebar.text_input("单号表路径", value=DEFAULT_ORDER_XLSX, key="p1")
        p_done = st.sidebar.text_input("直供订单表·处理完成（可选，1 个路径）", value=DEFAULT_PHOTO_XLSX, key="p_done")
        p_raw = st.sidebar.text_area("直供订单表·未处理（每行一个路径）", value="", key="p_raw", height=110)
        raw_paths = [ln.strip() for ln in (p_raw or "").splitlines() if ln.strip()]
        direct_paths = ([p_done] if p_done else []) + raw_paths
        all_paths = ([p1] if p1 else []) + direct_paths
        if p1 and direct_paths and all(os.path.exists(q) for q in all_paths):
            sheet1_path = p1
            done_files = [p_done] if p_done else []
            raw_files = raw_paths
        else:
            st.sidebar.warning("路径无效，请检查文件是否存在")
    else:
        f1 = st.sidebar.file_uploader("上传 单号表 (.xlsx)", type=["xlsx"], key="f1")
        f_done = st.sidebar.file_uploader("上传 直供订单表·处理完成（可选，1 个）", type=["xlsx"], key="f_done")
        f_raw = st.sidebar.file_uploader("上传 直供订单表·未处理（可多选）", type=["xlsx"],
                                         accept_multiple_files=True, key="f_raw")
        if f1 and (f_done or f_raw):
            uploaded_mode = True
            os.makedirs(UPLOAD_DIR, exist_ok=True)
            sheet1_path = os.path.join(UPLOAD_DIR, "order.xlsx")
            token = "|".join([_upload_token(f1)]
                             + ([_upload_token(f_done)] if f_done else [])
                             + [_upload_token(f) for f in f_raw])
            if st.session_state.uploaded_key != token:
                with open(sheet1_path, "wb") as fp:
                    fp.write(f1.getbuffer())
                done_files = [_save_upload(f_done, "done")] if f_done else []
                raw_files = [_save_upload(f, "raw") for f in f_raw]
                st.session_state.uploaded_key = token
                st.session_state.uploaded_paths = {"done": done_files, "raw": raw_files}
            else:
                saved = st.session_state.uploaded_paths
                done_files = saved.get("done", [])
                raw_files = saved.get("raw", [])
    st.sidebar.header("② 输出设置")
    export_dir = st.sidebar.text_input("图片保存目录", value=st.session_state.export_dir, key="export_dir_input")
    st.session_state.export_dir = (export_dir or EXPORT_DIR_DEFAULT).strip() or EXPORT_DIR_DEFAULT
    st.sidebar.caption("保存命名：订单编号_照片序号.扩展名")
    st.sidebar.button("解析并匹配数据", type="primary", width="stretch", key="btn_parse")
    return sheet1_path, done_files, raw_files, uploaded_mode


def do_parse(sheet1_path, done_files, raw_files):
    all_files = (done_files or []) + (raw_files or [])
    paths_ok = sheet1_path and all_files and all(
        q and os.path.exists(q) for q in [sheet1_path] + all_files)
    if not paths_ok:
        st.error("请先正确选择单号表，以及至少一个直供订单表文件（处理完成或未处理）")
        return False
    try:
        with st.spinner("正在解析单号表（订单编号）..."):
            order_list = xr.read_order_numbers(sheet1_path)
        merged_path = os.path.join(UPLOAD_DIR, "合并直供订单表.xlsx")
        with st.spinner("正在处理并拼接直供订单表（只保留 16 位三方单号且含照片的行）..."):
            result = xr.merge_direct_tables(all_files, set(order_list), merged_path)
    except Exception as e:
        st.error(f"解析失败：{e}")
        return False
    photo_orders = result["photo_orders"]
    seen = result["seen_orders"]
    with_photo = [o for o in order_list if o in photo_orders]
    no_photo = [o for o in order_list if o in (seen - set(photo_orders))]
    unmatched = [o for o in order_list if o not in seen]
    st.session_state.update(
        parse_done=True, sheet1_path=sheet1_path, merged_path=merged_path,
        orders=with_photo, order_urls=photo_orders,
        orders_no_photo=no_photo, orders_unmatched=unmatched,
        stats={"total": len(order_list), "with_photo": len(with_photo),
               "no_photo": len(no_photo), "unmatched": len(unmatched),
               "matched_rows": result["matched_rows"], "total_rows": result["total_rows"],
               "kept_rows": result["kept_rows"], "per_file": result["per_file"],
               "len_counter": result["len_counter"]},
        curr_index=0, curr_photo=0, marked=set(), saved={},
    )
    return True


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


@st.dialog("🖼️ 图片编辑（二级页面）", width="large")
def image_editor(order, ph, orig):
    """二级编辑页面：等比放大图片 + 旋转 + 微信截图式裁剪。"""
    pf = f"{order}|{ph}"
    st.markdown("**双击图片放大** · 按 **C** 键裁剪 · 拖拽画选区 · "
                "**Enter** 确认 / **Esc** 取消 / **方向键**微调 / 双击确认")
    rot = st.slider("旋转角度 (°)", -180, 180, st.session_state.get(f"rot_{pf}", 0), 1, key=f"drot_{pf}")
    rotated = it.rotate_image(orig, rot)
    reset_flag = st.session_state.crop_reset.get(pf, 0)
    crop_res = image_crop(img=_img_dataurl(rotated), width=720, dbl_opens=False,
                          reset=reset_flag, key=f"ecrop_{pf}")
    if isinstance(crop_res, dict):
        last_key = f"ecrop_res_{order}_{ph}"
        last_val = st.session_state.get(last_key)
        if crop_res != last_val:
            st.session_state[last_key] = crop_res
            if crop_res.get("confirmed") and isinstance(crop_res.get("crop"), dict):
                st.session_state.crop_box[pf] = crop_res["crop"]
                st.toast("✅ 已应用裁剪")
                st.rerun()
            if crop_res.get("canceled"):
                st.session_state.crop_box.pop(pf, None)
                st.toast("已取消裁剪")
                st.rerun()
    crop_box = st.session_state.crop_box.get(pf)
    processed = it.crop_image(rotated, crop_box) if crop_box else rotated
    c1, c2 = st.columns(2)
    c1.caption(f"编辑中（旋转 {rot}°）")
    c1.image(rotated, width="stretch")
    c2.caption("结果" + ("，已裁剪" if crop_box else ""))
    c2.image(processed, width="stretch")
    b1, b2, b3 = st.columns(3)
    if b1.button("✅ 应用并返回", type="primary", width="stretch"):
        st.session_state[f"rot_{pf}"] = rot
        st.session_state.editor_open = None
        st.rerun()
    if b2.button("↩️ 重置（旋转/裁剪）", width="stretch"):
        st.session_state[f"drot_{pf}"] = 0
        st.session_state.crop_box.pop(pf, None)
        st.session_state.crop_reset[pf] = reset_flag + 1
        st.rerun()
    if b3.button("✖️ 取消", width="stretch"):
        st.session_state.editor_open = None
        st.rerun()


def main():
    _init()
    sheet1_path, done_files, raw_files, uploaded_mode = sidebar()

    st.title("🖼️ 国补订单照片处理工具")
    st.caption("流程：选择单号表 + 直供订单表（处理完成1个/未处理多个）→ 处理拼接 → 匹配 → 逐单处理图片 → 保存 / 标记已上传")

    if st.session_state.btn_parse:
        if do_parse(sheet1_path, done_files, raw_files):
            st.session_state.uploaded_mode = uploaded_mode
            st.rerun()

    if not st.session_state.parse_done:
        st.info("👈 请在左侧选择单号表与直供订单表（处理完成 1 个 + 未处理可多选），然后点击“解析并匹配数据”。")
        return

    stats = st.session_state.stats
    m = st.columns(6)
    m[0].metric("单号表订单总数", stats.get("total", 0))
    m[1].metric("✅ 匹配且有照片", stats.get("with_photo", 0))
    m[2].metric("匹配但无照片", stats.get("no_photo", 0))
    m[3].metric("未匹配订单", stats.get("unmatched", 0))
    m[4].metric("直供表总行数", stats.get("total_rows", 0))
    m[5].metric("生成文件行数", stats.get("kept_rows", 0))
    st.caption(f"单号表：{st.session_state.sheet1_path}　|　图片输出：{st.session_state.export_dir}")

    merged_path = st.session_state.get("merged_path")
    if merged_path and os.path.exists(merged_path):
        with open(merged_path, "rb") as fp:
            st.download_button("⬇️ 下载生成的新文件（合并直供订单表.xlsx）", data=fp.read(),
                               file_name=os.path.basename(merged_path),
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               key="dl_merged")
        st.caption(f"📦 生成文件位置：{merged_path}")

    per_file = stats.get("per_file") or []
    if per_file:
        with st.expander(f"各直供表文件处理统计（{len(per_file)} 个）"):
            for pf in per_file:
                st.write(f"📄 {pf['name']}：原始 {pf['total']} 行 → 保留（16位且有照片）{pf['kept']} 行")

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
    st.subheader("🛠️ 图片处理")
    st.caption("**双击左侧预览图**进入二级编辑页（等比放大 + 旋转 + 微信截图式裁剪）")

    # 主预览组件：双击 → 打开二级编辑页
    prev_res = image_crop(img=_img_dataurl(orig), width=460, dbl_opens=True, key=f"prev_{pf}")
    if isinstance(prev_res, dict) and prev_res.get("open_editor"):
        st.session_state.editor_open = pf
        st.rerun()

    rot = st.session_state.get(f"rot_{pf}", 0)
    crop_box = st.session_state.crop_box.get(pf)
    reset_flag = st.session_state.crop_reset.get(pf, 0)
    rotated = it.rotate_image(orig, rot)
    processed = it.crop_image(rotated, crop_box) if crop_box else rotated

    c1, c2 = st.columns(2)
    c1.caption("原图")
    c1.image(orig, width="stretch")
    c2.caption("当前结果" + (f"（旋转 {rot}°" + ("，已裁剪" if crop_box else "") + "）" if (rot or crop_box) else ""))
    c2.image(processed, width="stretch")

    st.divider()
    done = (order, ph) in st.session_state.saved
    bc = st.columns([2, 1, 1, 1])
    auto_next = bc[0].checkbox("保存后自动跳到下一张/下一单", value=True, key="auto_next")
    if bc[1].button("💾 保存当前图片" + ("（已保存·可覆盖）" if done else ""), type="primary", width="stretch"):
        outdir = st.session_state.export_dir
        os.makedirs(outdir, exist_ok=True)
        name = f"{order}_{ph + 1}.jpg"
        path = os.path.join(outdir, name)
        it.save_image(processed, path, "JPG", 90)
        st.session_state.saved[(order, ph)] = path
        st.toast(f"✅ 已保存：{name}")
        if auto_next:
            advance_after_save(len(urls), len(orders), idx)
            st.rerun()
    if bc[2].button("↩️ 重置本图（旋转/裁剪）", width="stretch"):
        st.session_state[f"rot_{pf}"] = 0
        st.session_state.crop_box.pop(pf, None)
        st.session_state.crop_reset[pf] = reset_flag + 1
        st.rerun()
    if bc[3].button("⬇️ 下载此图", width="stretch"):
        st.download_button("点击下载", data=it.image_to_bytes(processed, "JPG", 90),
                           file_name=f"{order}_{ph + 1}.jpg", key=f"dl_{pf}")

    # 二级编辑页面（双击预览图后弹出）
    if st.session_state.get("editor_open") == pf:
        image_editor(order, ph, orig)

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
    try:
        main()
    except Exception as e:
        import traceback
        tb = traceback.format_exc()
        try:
            err_path = os.path.join(UPLOAD_DIR, "error.log")
            with open(err_path, "w", encoding="utf-8") as fp:
                fp.write(tb)
        except Exception:
            pass
        st.error(f"程序发生错误：{type(e).__name__}: {e}\n\n完整错误已写入 {UPLOAD_DIR}\\error.log，请把日志内容反馈给开发者。")


