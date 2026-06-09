"""
自定义 GUI 控件 — TwoWayScrollableFrame

基于 Canvas 的双向（垂直+水平）可滚动容器，
内部嵌入一个 CTkFrame 用于放置子控件。
"""

import customtkinter as ctk


class TwoWayScrollableFrame(ctk.CTkFrame):
    """双向滚动框架（垂直 + 水平），支持鼠标滚轮。"""

    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self.canvas_bg = self._get_theme_bg()
        self.canvas = ctk.CTkCanvas(self, highlightthickness=0, bd=0, bg=self.canvas_bg)
        self.canvas.grid(row=0, column=0, sticky="nsew")

        self.v_scrollbar = ctk.CTkScrollbar(self, orientation="vertical", command=self.canvas.yview)
        self.v_scrollbar.grid(row=0, column=1, sticky="ns")

        self.h_scrollbar = ctk.CTkScrollbar(self, orientation="horizontal", command=self.canvas.xview)
        self.h_scrollbar.grid(row=1, column=0, sticky="ew")

        self.canvas.configure(yscrollcommand=self.v_scrollbar.set, xscrollcommand=self.h_scrollbar.set)

        self.inner_frame = ctk.CTkFrame(self.canvas, fg_color="transparent")
        self.canvas_window = self.canvas.create_window((0, 0), window=self.inner_frame, anchor="nw")

        self.inner_frame.bind("<Configure>", self._on_inner_frame_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)

        for w in (self.canvas, self.inner_frame):
            w.bind("<MouseWheel>", self._on_mousewheel)
            w.bind("<Button-4>", self._on_mousewheel_linux)
            w.bind("<Button-5>", self._on_mousewheel_linux)

    def _get_theme_bg(self):
        if ctk.get_appearance_mode() == "light":
            return "#F8FAFC"
        else:
            return "#0F172A"

    def update_theme_colors(self):
        new_bg = self._get_theme_bg()
        self.canvas.configure(bg=new_bg)

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _on_mousewheel_linux(self, event):
        if event.num == 4:
            self.canvas.yview_scroll(-1, "units")
        elif event.num == 5:
            self.canvas.yview_scroll(1, "units")

    def _on_inner_frame_configure(self, event):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event):
        canvas_width = event.width
        self.canvas.itemconfig(self.canvas_window, width=max(canvas_width, self.inner_frame.winfo_reqwidth()))

    def get_inner_frame(self):
        return self.inner_frame
