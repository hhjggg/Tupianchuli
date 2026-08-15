# -*- coding: utf-8 -*-
"""国补订单照片处理工具 - Streamlit 主程序

运行: cd deepseek && streamlit run app.py
"""
import base64
import glob
import io
import os
import sys
import threading
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


def _thumb_dataurl(img, max_w=240):
    """缩略图（等比缩小后）→ base64 data URL。"""
    if img.width > max_w:
        r = max_w / img.width
        img = img.resize((max_w, max(1, int(img.height * r))), Image.LANCZOS)
    return _img_dataurl(img)


def _img_dataurl_small(img, max_w=900):
    """缩小后转 base64（供组件显示，大幅提升渲染速度；下载仍用全分辨率）。"""
    if img.width > max_w:
        r = max_w / img.width
        img = img.resize((max_w, max(1, int(img.height * r))), Image.LANCZOS)
    return _img_dataurl(img)


def _init():
    st.session_state.setdefault("parse_done", False)
    st.session_state.setdefault("orders", [])
    st.session_state.setdefault("orders_all", [])
    st.session_state.setdefault("order_urls", {})
    st.session_state.setdefault("orders_no_photo", [])
    st.session_state.setdefault("orders_unmatched", [])
    st.session_state.setdefault("stats", {})
    st.session_state.setdefault("curr_index", 0)
    st.session_state.setdefault("curr_photo", 0)
    st.session_state.setdefault("view_order", None)
    st.session_state.setdefault("marked", set())
    st.session_state.setdefault("upload_status", {})
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
    st.session_state.setdefault("proc", {})
    st.session_state.setdefault("rot_state", {})
    st.session_state.setdefault("prefetched", set())
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
        orders=with_photo, orders_all=order_list, order_urls=photo_orders,
        orders_no_photo=no_photo, orders_unmatched=unmatched,
        stats={"total": len(order_list), "with_photo": len(with_photo),
               "no_photo": len(no_photo), "unmatched": len(unmatched),
               "matched_rows": result["matched_rows"], "total_rows": result["total_rows"],
               "kept_rows": result["kept_rows"], "per_file": result["per_file"],
               "len_counter": result["len_counter"]},
        curr_index=0, curr_photo=0, view_order=None, marked=set(), saved={}, upload_status={},
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


def _prefetch_urls(urls):
    """后台线程：把照片 URL 逐个下载到本地缓存（纯函数，不访问 session_state）。"""
    for u in urls:
        it.ensure_downloaded(u, PHOTO_CACHE)


def _trigger_prefetch():
    """预下载下一个订单的照片到缓存（后台线程，不阻塞当前页面）。"""
    orders = st.session_state.get("orders", [])
    idx = st.session_state.get("curr_index", 0)
    if orders and idx + 1 < len(orders):
        nxt = orders[idx + 1]
        st.session_state.setdefault("prefetched", set())
        if nxt not in st.session_state["prefetched"]:
            st.session_state["prefetched"].add(nxt)
            urls = st.session_state.get("order_urls", {}).get(nxt, [])
            if urls:
                threading.Thread(target=_prefetch_urls, args=(urls,), daemon=True).start()


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


def _reset_editor_cb(order, ph, orig):
    """重置二级编辑页（on_click 回调，在 widget 实例化前执行）。"""
    pf = f"{order}|{ph}"
    st.session_state.setdefault("proc", {})
    st.session_state.setdefault("rot_state", {})
    st.session_state.setdefault("crop_reset", {})
    st.session_state["proc"][pf] = orig.copy()
    st.session_state["rot_state"][pf] = 0
    st.session_state[f"drot_{pf}"] = 0
    st.session_state.pop(f"ecrop_res_{order}_{ph}", None)
    st.session_state["crop_reset"][pf] = st.session_state["crop_reset"].get(pf, 0) + 1


def _prev_order_cb(idx, n):
    """上一单（on_click 回调，在 widget 实例化前执行，可同步 selectbox 值）。"""
    nidx = max(0, idx - 1)
    st.session_state.curr_index = nidx
    st.session_state.curr_photo = 0
    st.session_state["order_jump"] = nidx


def _next_order_cb(idx, n):
    """下一单（on_click 回调）。"""
    nidx = min(n - 1, idx + 1)
    st.session_state.curr_index = nidx
    st.session_state.curr_photo = 0
    st.session_state["order_jump"] = nidx


def render_upload_status(order):
    """渲染“上传状态”标记区（有照片/无照片订单通用）。

    使用 st.form：选择上传状态不会触发页面刷新，点「标记」才提交。
    """
    st.divider()
    marked = order in st.session_state.get("marked", set())
    marked_val = st.session_state.get("upload_status", {}).get(order, "")
    status_text = ("✅ 已标记：" + marked_val) if marked else "⬜ 未标记"
    st.markdown(f"**上传状态：** {status_text}")
    with st.form(key=f"up_form_{order}"):
        col1, col2 = st.columns([2, 1])
        col1.selectbox("上传状态", ["是", "无上传通道"], key=f"up_status_{order}")
        submitted = col2.form_submit_button("📤 标记")
    if submitted:
        up_status = st.session_state.get(f"up_status_{order}", "是")
        path1 = st.session_state.sheet1_path
        try:
            with st.spinner("正在写入单号表..."):
                n = xr.mark_uploaded(path1, order, up_status)
            if n:
                st.session_state.setdefault("marked", set()).add(order)
                st.session_state.setdefault("upload_status", {})[order] = up_status
                st.success(f"已将订单 {order} 的“是否上传”列写入“{up_status}”（共更新 {n} 行）")
                if st.session_state.uploaded_mode:
                    with open(path1, "rb") as fp:
                        st.download_button("⬇️ 下载已更新的单号表", data=fp.read(),
                                           file_name=os.path.basename(path1),
                                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            else:
                st.warning(f"在单号表中未找到订单编号 {order}，未写入")
        except Exception as e:
            st.error(f"写入失败：{e}")


def _jump_select_cb():
    """跳转到订单下拉框 on_change 回调（在 widget 实例化前执行，可安全设置状态）。"""
    sel = st.session_state.get("order_jump")
    orders = st.session_state.get("orders", [])
    cur = st.session_state.get("current_order", "")
    if sel and sel in orders:
        idx = orders.index(sel)
        st.session_state.curr_index = idx
        st.session_state.curr_photo = 0
        st.session_state.editor_open = None
        st.session_state["view_order"] = None
        st.session_state[f"photo_sel_{sel}"] = 0
    else:
        orders_all = st.session_state.get("orders_all", orders)
        if sel in orders_all:
            st.session_state["view_order"] = sel  # 无照片详情视图（可标记上传）
            st.session_state.editor_open = None
        else:
            st.session_state["jump_msg"] = f"⚠️ 订单号 {sel or '（空）'} 不存在"
            if cur:
                st.session_state["order_jump"] = cur  # 恢复下拉框为当前订单


def _jump_to_order():
    """按订单号直达（on_click 回调，在 widget 实例化前执行）。"""
    target = (st.session_state.get("jump_order_no") or "").strip()
    orders = st.session_state.get("orders", [])
    orders_all = st.session_state.get("orders_all", orders)
    if target and target in orders:
        idx = orders.index(target)
        st.session_state.curr_index = idx
        st.session_state.curr_photo = 0
        st.session_state["order_jump"] = target
        st.session_state.editor_open = None
        st.session_state["view_order"] = None
        st.session_state[f"photo_sel_{target}"] = 0
        st.session_state["jump_msg"] = f"✅ 已跳转到订单 {target}"
    elif target in orders_all:
        st.session_state["view_order"] = target  # 无照片详情视图（可标记上传）
        st.session_state.editor_open = None
        st.session_state["jump_msg"] = f"📄 已查看订单 {target}（无照片，可标记上传状态）"
    else:
        st.session_state["jump_msg"] = f"⚠️ 订单号 {target or '（空）'} 不存在"


def _fname_change_cb(pf):
    """文件名输入框值变化时，立即提交到独立键（on_change 回调）。"""
    st.session_state[f"fname_val_{pf}"] = st.session_state.get(f"fname_{pf}", "").strip()


@st.dialog("🖼️ 图片编辑（二级页面）", width="large")
def image_editor(order, ph, orig):
    """二级编辑页面：所见即所得 —— 旋转/裁剪直接作用并显示在同一张图上。"""
    # 从 session_state 解析最新编辑目标（dialog fragment 重跑时传入参数可能为旧值）
    cur_open = st.session_state.get("editor_open")
    if cur_open:
        try:
            cur_o, cur_i_s = str(cur_open).split("|", 1)
            cur_i = int(cur_i_s)
            if cur_o == str(order) and cur_i != ph:
                ph = cur_i
                new_orig = get_original(order, ph)
                if new_orig is not None:
                    orig = new_orig
        except (ValueError, AttributeError):
            pass
    pf = f"{order}|{ph}"
    # 确保会话状态键存在（dialog fragment 可能独立执行）
    st.session_state.setdefault("crop_box", {})
    st.session_state.setdefault("crop_reset", {})
    st.session_state.setdefault("editor_open", None)
    st.session_state.setdefault("proc", {})
    st.session_state.setdefault("rot_state", {})
    if pf not in st.session_state["proc"]:
        st.session_state["proc"][pf] = orig.copy()
    st.markdown("**双击图片放大** · 按 **X** 键裁剪 · 拖拽画选区 · "
                "**Enter** 确认 / **Esc** 取消 / **方向键**微调 / 双击确认")
    prev_rot = st.session_state.get("rot_state", {}).get(pf, 0)
    # 处理重置标志（在 slider 实例化前同步 drot，避免 widget 修改报错）
    if st.session_state.pop(f"reset_flag_{pf}", False):
        st.session_state[f"drot_{pf}"] = 0
    st.session_state.setdefault(f"drot_{pf}", prev_rot)
    rot = st.slider("旋转角度 (°)", -180, 180, step=90, key=f"drot_{pf}")
    # 旋转：增量应用到当前图（效果直接显示）
    cur = st.session_state["proc"][pf]
    delta = rot - prev_rot
    if delta:
        cur = it.rotate_image(cur, delta)
        st.session_state["proc"][pf] = cur
        st.session_state["rot_state"][pf] = rot
    reset_flag = st.session_state.get("crop_reset", {}).get(pf, 0)
    # 上一张 / 图片 / 下一张（按钮在左右两侧、垂直居中；直接改状态并完整刷新）
    urls_len = len(st.session_state.get("order_urls", {}).get(str(order), [ph]))
    nav = st.columns([1, 8, 1], vertical_alignment="center")
    if nav[0].button("◀ 上一张", key=f"eprev_{pf}", disabled=(ph <= 0), width="stretch"):
        new_idx = max(0, ph - 1)
        st.session_state.editor_open = f"{order}|{new_idx}"
        st.session_state.curr_photo = new_idx
        st.session_state.pop(f"ecrop_res_{order}_{new_idx}", None)
    with nav[1]:
        crop_res = image_crop(img=_img_dataurl_small(cur), width=720, dbl_opens=False,
                              reset=reset_flag, key=f"ecrop_{pf}")
    if nav[2].button("下一张 ▶", key=f"enext_{pf}", disabled=(ph >= urls_len - 1), width="stretch"):
        new_idx = min(urls_len - 1, ph + 1)
        st.session_state.editor_open = f"{order}|{new_idx}"
        st.session_state.curr_photo = new_idx
        st.session_state.pop(f"ecrop_res_{order}_{new_idx}", None)
    if isinstance(crop_res, dict):
        last_key = f"ecrop_res_{order}_{ph}"
        last_val = st.session_state.get(last_key)
        if crop_res != last_val:
            st.session_state[last_key] = crop_res
            if crop_res.get("confirmed") and isinstance(crop_res.get("crop"), dict):
                crop = crop_res["crop"]
                # 组件显示的是缩小图(max_w=900)，裁剪坐标需换算回全尺寸
                ratio = cur.width / min(cur.width, 900)
                crop_full = {
                    "x": int(crop.get("x", 0) * ratio),
                    "y": int(crop.get("y", 0) * ratio),
                    "w": int(crop.get("w", 0) * ratio),
                    "h": int(crop.get("h", 0) * ratio),
                }
                st.session_state["proc"][pf] = it.crop_image(cur, crop_full)
                # 清除该订单缩略图缓存，确保主页面缩略图始终从原始文件显示原图
                for _ti in range(len(st.session_state.get("order_urls", {}).get(str(order), []))):
                    st.session_state.pop(f"thumb_{order}_{_ti}", None)
                st.toast("✅ 已应用裁剪")
                # 不调用 st.rerun()：组件交互会自动触发 dialog 重渲染，保持对话框打开
            if crop_res.get("canceled"):
                st.toast("已取消裁剪")
    display = st.session_state["proc"][pf]
    # 命名 + 格式
    fcol1, fcol2 = st.columns(2)
    fcol1.text_input("文件名（不含扩展名）", value=f"{order}_{ph + 1}", key=f"fname_{pf}",
                     on_change=_fname_change_cb, args=(pf,))
    fmt_ = fcol2.selectbox("保存格式", ["PNG", "JPG", "WEBP"], index=0, key=f"fmt_{pf}")
    fname_s = (st.session_state.get(f"fname_val_{pf}") or st.session_state.get(f"fname_{pf}")
               or f"{order}_{ph + 1}").strip() or f"{order}_{ph + 1}"
    ext = it.FORMAT_EXT.get(fmt_, "png")
    b1, b2 = st.columns(2)
    if sys.platform == "win32":
        # 本地运行：直接保存到桌面（不弹系统窗口）
        if b1.button("💾 保存到桌面", type="primary", width="stretch"):
            outdir = st.session_state.export_dir or EXPORT_DIR_DEFAULT
            os.makedirs(outdir, exist_ok=True)
            path = os.path.join(outdir, f"{fname_s}.{ext}")
            it.save_image(display, path, fmt_, 95)
            st.session_state.saved[(order, ph)] = f"{fname_s}.{ext}"
            st.toast(f"✅ 已保存：{path}")
            st.rerun()
    else:
        dl = b1.download_button("💾 下载处理图（保存到本地）",
                                data=it.image_to_bytes(display, fmt_, 95),
                                file_name=f"{fname_s}.{ext}", type="primary", key=f"dl_{pf}")
        if dl:
            st.session_state.saved[(order, ph)] = f"{fname_s}.{ext}"
            st.toast(f"✅ 已下载：{fname_s}.{ext}")
            st.rerun()
    if b2.button("✖️ 关闭", width="stretch"):
        st.session_state.editor_open = None
        st.session_state.pop(f"dialog_trigger_{pf}", None)
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

    orders = st.session_state.orders
    # 无照片/未匹配订单详情视图（仅标记上传状态）
    view_order = st.session_state.get("view_order")
    if view_order and view_order not in orders:
        st.subheader(f"📄 订单 {view_order}")
        st.info("该订单无照片或未匹配，无法进行图片处理，仅可标记上传状态。")
        render_upload_status(view_order)
        if st.button("← 返回可处理订单", width="stretch"):
            st.session_state["view_order"] = None
            if orders:
                st.session_state["order_jump"] = orders[min(st.session_state.get("curr_index", 0), len(orders) - 1)]
            st.rerun()
        return
    if not orders:
        st.warning("没有找到任何匹配且有照片的订单，请检查两张数据表是否选择正确。")
        return
    idx = min(st.session_state.curr_index, len(orders) - 1)
    st.session_state.curr_index = idx
    order = orders[idx]
    urls = st.session_state.order_urls[order]

    st.divider()
    hc = st.columns([4, 1, 1])
    all_orders = st.session_state.get("orders_all") or orders
    cur_all_pos = all_orders.index(order) + 1 if order in all_orders else idx + 1
    hc[0].markdown(f"### 📋 订单 {cur_all_pos} / {len(all_orders)}")
    hc[0].caption("当前订单号（可复制）：")
    hc[0].code(str(order))
    if hc[1].button("‹ 上一单", key="prev_order", width="stretch"):
        nidx = max(0, idx - 1)
        st.session_state.curr_index = nidx
        st.session_state.curr_photo = 0
        st.session_state["view_order"] = None
        st.session_state["order_jump"] = orders[nidx]  # 同步下拉框（订单号）
        st.rerun()
    if hc[2].button("下一单 ›", key="next_order", width="stretch"):
        nidx = min(len(orders) - 1, idx + 1)
        st.session_state.curr_index = nidx
        st.session_state.curr_photo = 0
        st.session_state["view_order"] = None
        st.session_state["order_jump"] = orders[nidx]
        st.rerun()

    # 跳转到订单：显示全部订单（不管匹配/照片状态）
    all_orders = st.session_state.get("orders_all") or orders
    cur_all_idx = all_orders.index(order) if order in all_orders else 0
    _urls = st.session_state.order_urls
    _np_set = set(st.session_state.get("orders_no_photo", []))
    _um_set = set(st.session_state.get("orders_unmatched", []))
    _p_set = set(orders)
    st.session_state["current_order"] = order
    sel = st.selectbox("跳转到订单（全部订单）", all_orders, index=cur_all_idx,
                       format_func=lambda o, _p=_p_set, _np=_np_set, _um=_um_set, _u=_urls:
                       (f"{o}（{len(_u[o])}张）" if o in _p else
                        (f"{o}（无照片）" if o in _np else f"{o}（未匹配）")),
                       key="order_jump", on_change=_jump_select_cb)

    # 订单号直达（可复制/粘贴订单号，存在则跳转，不存在则提示）
    jc = st.columns([3, 1])
    jc[0].text_input("订单号直达（输入或粘贴订单号，点「跳转」）", key="jump_order_no")
    if jc[1].button("🔍 跳转", width="stretch", on_click=_jump_to_order):
        pass
    jmsg = st.session_state.get("jump_msg")
    if jmsg:
        if jmsg.startswith("✅"):
            st.success(jmsg)
        else:
            st.warning(jmsg)



    ph = min(st.session_state.curr_photo, len(urls) - 1)
    st.session_state.curr_photo = ph
    st.markdown(f"#### 🖼️ 该订单共 {len(urls)} 张照片")

    thumbs = get_order_photos(order)
    _trigger_prefetch()  # 后台预下载下一个订单的照片
    tcols = st.columns(min(len(urls), 6))
    for i, u in enumerate(urls):
        path, ok = thumbs.get(u, (None, False))
        with tcols[i % len(tcols)]:
            if ok:
                cap = f"第{i+1}张" + (" ✅已存" if (order, i) in st.session_state.saved else "")
                tb_key = f"thumb_{order}_{i}"
                if tb_key not in st.session_state:
                    t_img = Image.open(path).convert("RGB")
                    st.session_state[tb_key] = _thumb_dataurl(t_img)
                th_res = image_crop(img=st.session_state[tb_key], width=150, dbl_opens=True, key=f"th_{order}_{i}")
                if isinstance(th_res, dict) and th_res.get("open_editor"):
                    ts = th_res.get("open_ts")
                    last_key = f"th_open_{order}_{i}"
                    if ts is not None and ts != st.session_state.get(last_key):
                        st.session_state[last_key] = ts
                        st.session_state.editor_open = f"{order}|{i}"
                        st.session_state[f"dialog_trigger_{order}|{i}"] = True  # 一次性打开标志
                        st.session_state.curr_photo = i
                        st.session_state[f"photo_sel_{order}"] = i  # 同步照片选择，避免残留覆盖
                        # 打开编辑页时清除处理副本，确保每次从原始留底重新开始
                        st.session_state["proc"].pop(f"{order}|{i}", None)
                        st.session_state["rot_state"].pop(f"{order}|{i}", None)
                        st.session_state[f"reset_flag_{order}|{i}"] = True
                        st.rerun()
                st.caption(cap)
            else:
                st.warning(f"第{i+1}张下载失败")

    _saved = st.session_state.saved
    psel = st.radio("选择要处理的照片", range(len(urls)), index=ph,
                    format_func=lambda i, _order=order, _saved=_saved: f"第 {i+1} 张" + ("（已保存）" if (_order, i) in _saved else ""),
                    horizontal=True, key=f"photo_sel_{order}")
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

    st.subheader("🛠️ 图片处理")
    st.caption("**双击上方照片缩略图**进入二级编辑（旋转 + 微信截图式裁剪 + 命名 + 下载）")

    # 二级编辑页面（双击照片缩略图后弹出，一次性触发）
    edit_key = st.session_state.get("editor_open")
    if edit_key:
        e_order = e_idx = None
        try:
            e_order, e_idx_s = str(edit_key).split("|", 1)
            e_idx = int(e_idx_s)
        except (ValueError, AttributeError):
            pass
        if e_order == str(order) and e_idx is not None and 0 <= e_idx < len(urls):
            if st.session_state.pop(f"dialog_trigger_{edit_key}", False):
                # 双击触发的打开：真正打开 dialog
                e_orig = get_original(order, e_idx)
                if e_orig is not None:
                    image_editor(order, e_idx, e_orig)
            else:
                # 无触发标志：dialog 已被关闭（如右上角 X），清除残留，避免主页面交互重新弹出
                st.session_state.pop("editor_open", None)

    render_upload_status(order)


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


