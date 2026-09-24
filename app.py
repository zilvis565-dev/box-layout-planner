from flask import Flask, render_template, request, redirect, url_for, send_file, session
import sqlite3
from datetime import datetime
from openpyxl import Workbook
import plotly.graph_objects as go
import plotly.io as pio
import math
import copy

app = Flask(__name__)
app.secret_key = "box-layout-secret"

DB_NAME = "/tmp/planner.db"

BOX_COLORS = [
    "#3b82f6", "#22c55e", "#f59e0b", "#ef4444",
    "#8b5cf6", "#06b6d4", "#eab308", "#14b8a6",
    "#f97316", "#84cc16"
]

DEFAULT_PALLETS = [
    {"name": "EUR 1200 x 800", "length": 1200, "width": 800, "height": 150, "max_weight": 500},
    {"name": "FIN 1200 x 1000", "length": 1200, "width": 1000, "height": 150, "max_weight": 700},
    {"name": "US 1219 x 1016", "length": 1219, "width": 1016, "height": 150, "max_weight": 700},
    {"name": "Half pallet 800 x 600", "length": 800, "width": 600, "height": 150, "max_weight": 300},
    {"name": "1200 x 1050", "length": 1200, "width": 1050, "height": 150, "max_weight": 700},
    {"name": "1200 x 950", "length": 1200, "width": 950, "height": 150, "max_weight": 700},
    {"name": "1500 x 1100", "length": 1500, "width": 1100, "height": 150, "max_weight": 1000},
]

DEFAULT_BOXES = [
    {"name": "1160x450x450", "length": 1160, "width": 450, "height": 450, "weight": 1.0},
    {"name": "710x430x270", "length": 710, "width": 430, "height": 270, "weight": 1.0},
    {"name": "1210x450x330", "length": 1210, "width": 450, "height": 330, "weight": 1.0},
    {"name": "800x470x290", "length": 800, "width": 470, "height": 290, "weight": 1.0},
]

LOADING_TARGETS = [
    "Sea container",
    "Truck",
    "Trailer",
    "Van",
    "Warehouse storage",
    "Custom"
]


def get_conn():
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_conn()
    c = conn.cursor()

    c.execute("""
        CREATE TABLE IF NOT EXISTS box_library (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            length INTEGER NOT NULL,
            width INTEGER NOT NULL,
            height INTEGER NOT NULL,
            weight REAL NOT NULL
        )
    """)

    c.execute("""
        CREATE TABLE IF NOT EXISTS pallet_library (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            length INTEGER NOT NULL,
            width INTEGER NOT NULL,
            height INTEGER NOT NULL,
            max_weight REAL NOT NULL
        )
    """)

    c.execute("""
        CREATE TABLE IF NOT EXISTS plans (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            plan_name TEXT NOT NULL,
            loading_target TEXT NOT NULL,
            max_load_height INTEGER NOT NULL,
            overhang INTEGER NOT NULL,
            allow_rotation INTEGER NOT NULL DEFAULT 1,
            created_at TEXT NOT NULL
        )
    """)

    c.execute("""
        CREATE TABLE IF NOT EXISTS plan_boxes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            plan_id INTEGER NOT NULL,
            box_name TEXT NOT NULL,
            length INTEGER NOT NULL,
            width INTEGER NOT NULL,
            height INTEGER NOT NULL,
            weight REAL NOT NULL,
            qty INTEGER NOT NULL,
            FOREIGN KEY(plan_id) REFERENCES plans(id)
        )
    """)

    c.execute("""
        CREATE TABLE IF NOT EXISTS plan_allowed_pallets (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            plan_id INTEGER NOT NULL,
            pallet_library_id INTEGER NOT NULL,
            FOREIGN KEY(plan_id) REFERENCES plans(id),
            FOREIGN KEY(pallet_library_id) REFERENCES pallet_library(id)
        )
    """)

    for pallet in DEFAULT_PALLETS:
        existing = c.execute("SELECT id FROM pallet_library WHERE name = ?", (pallet["name"],)).fetchone()
        if not existing:
            c.execute("""
                INSERT INTO pallet_library (name, length, width, height, max_weight)
                VALUES (?, ?, ?, ?, ?)
            """, (
                pallet["name"],
                pallet["length"],
                pallet["width"],
                pallet["height"],
                pallet["max_weight"]
            ))

    for box in DEFAULT_BOXES:
        existing = c.execute("SELECT id FROM box_library WHERE name = ?", (box["name"],)).fetchone()
        if not existing:
            c.execute("""
                INSERT INTO box_library (name, length, width, height, weight)
                VALUES (?, ?, ?, ?, ?)
            """, (
                box["name"],
                box["length"],
                box["width"],
                box["height"],
                box["weight"]
            ))

    conn.commit()
    conn.close()


def build_plan_from_pallet(base_plan, pallet_row):
    return {
        "pallet_type": pallet_row["name"],
        "pallet_length": pallet_row["length"],
        "pallet_width": pallet_row["width"],
        "pallet_height": pallet_row["height"],
        "max_load_height": base_plan["max_load_height"],
        "max_pallet_weight": pallet_row["max_weight"],
        "overhang": base_plan["overhang"],
        "allow_rotation": base_plan["allow_rotation"],
        "loading_target": base_plan["loading_target"]
    }


def merge_same_box_type_rows(boxes):
    if not boxes:
        return None

    return {
        "box_name": boxes[0]["box_name"],
        "length": boxes[0]["length"],
        "width": boxes[0]["width"],
        "height": boxes[0]["height"],
        "weight": boxes[0]["weight"],
        "qty": sum(box["qty"] for box in boxes)
    }


def single_box_capacity(base_plan, pallet_row, box):
    plan = build_plan_from_pallet(base_plan, pallet_row)

    effective_length = plan["pallet_length"] + 2 * plan["overhang"]
    effective_width = plan["pallet_width"] + 2 * plan["overhang"]
    usable_height = plan["max_load_height"] - plan["pallet_height"]

    orientations = [(box["length"], box["width"])]
    if plan["allow_rotation"] and box["length"] != box["width"]:
        orientations.append((box["width"], box["length"]))

    best = None

    for bl, bw in orientations:
        per_row = effective_length // bl
        per_col = effective_width // bw
        per_layer = per_row * per_col
        layers = usable_height // box["height"] if box["height"] > 0 else 0

        if per_row <= 0 or per_col <= 0 or per_layer <= 0 or layers <= 0:
            continue

        by_height = per_layer * layers
        by_weight = int(plan["max_pallet_weight"] // box["weight"]) if box["weight"] > 0 else by_height
        per_pallet = min(by_height, by_weight)

        if per_pallet <= 0:
            continue

        candidate = {
            "orientation": (bl, bw),
            "per_row": int(per_row),
            "per_col": int(per_col),
            "per_layer": int(per_layer),
            "layers": int(layers),
            "per_pallet": int(per_pallet),
            "effective_length": int(effective_length),
            "effective_width": int(effective_width),
            "usable_height": int(usable_height),
            "pallet_type": pallet_row["name"]
        }

        if best is None or candidate["per_pallet"] > best["per_pallet"]:
            best = candidate

    return best


def solve_single_box_type_strict(base_plan, allowed_pallets, merged_box):
    best_result = None

    for pallet_row in allowed_pallets:
        cap = single_box_capacity(base_plan, pallet_row, merged_box)
        if not cap:
            continue

        plan = build_plan_from_pallet(base_plan, pallet_row)
        qty = merged_box["qty"]
        pallets_needed = math.ceil(qty / cap["per_pallet"])

        pallets = []
        remaining = qty

        for _ in range(pallets_needed):
            pallet_qty = min(remaining, cap["per_pallet"])
            remaining -= pallet_qty

            pallet = {
                "boxes": [],
                "weight": pallet_qty * merged_box["weight"],
                "used_height": plan["pallet_height"],
                "layers": [],
                "pallet_type": plan["pallet_type"],
                "pallet_length": plan["pallet_length"],
                "pallet_width": plan["pallet_width"],
                "pallet_height": plan["pallet_height"],
                "max_weight": plan["max_pallet_weight"]
            }

            bl, bw = cap["orientation"]
            z = plan["pallet_height"]
            qty_left = pallet_qty

            for _layer in range(cap["layers"]):
                if qty_left <= 0:
                    break

                for row in range(cap["per_col"]):
                    for col in range(cap["per_row"]):
                        if qty_left <= 0:
                            break

                        x = col * bl
                        y = row * bw

                        pallet["boxes"].append({
                            "box_name": merged_box["box_name"],
                            "length": merged_box["length"],
                            "width": merged_box["width"],
                            "height": merged_box["height"],
                            "weight": merged_box["weight"],
                            "placed_length": bl,
                            "placed_width": bw,
                            "x": x,
                            "y": y,
                            "z": z
                        })
                        qty_left -= 1

                pallet["used_height"] = max(pallet["used_height"], z + merged_box["height"])
                z += merged_box["height"]

            pallets.append(pallet)

        result = {
            "plan": plan,
            "pallets": pallets,
            "debug": cap
        }

        if best_result is None:
            best_result = result
        else:
            current_count = len(result["pallets"])
            best_count = len(best_result["pallets"])
            if current_count < best_count:
                best_result = result
            elif current_count == best_count and cap["per_pallet"] > best_result["debug"]["per_pallet"]:
                best_result = result

    return best_result


def solve_multi_box_simple(base_plan, allowed_pallets, boxes):
    pallets = []

    grouped = {}
    for box in boxes:
        key = box["box_name"]
        if key not in grouped:
            grouped[key] = {
                "box_name": box["box_name"],
                "length": box["length"],
                "width": box["width"],
                "height": box["height"],
                "weight": box["weight"],
                "qty": 0
            }
        grouped[key]["qty"] += box["qty"]

    for box_name in grouped:
        result = solve_single_box_type_strict(base_plan, allowed_pallets, grouped[box_name])
        if result:
            pallets.extend(result["pallets"])

    return pallets


def base_and_volume_utilization(pallet, plan):
    pallet_area = (pallet["pallet_length"] + 2 * plan["overhang"]) * (pallet["pallet_width"] + 2 * plan["overhang"])
    base_z = min(box["z"] for box in pallet["boxes"]) if pallet["boxes"] else 0
    base_used_area = sum(
        box["placed_length"] * box["placed_width"]
        for box in pallet["boxes"]
        if box["z"] == base_z
    )
    base_util = round((base_used_area / pallet_area * 100), 2) if pallet_area > 0 else 0

    pallet_volume = pallet_area * plan["max_load_height"]
    used_volume = sum(
        box["placed_length"] * box["placed_width"] * box["height"]
        for box in pallet["boxes"]
    )
    volume_util = round((used_volume / pallet_volume * 100), 2) if pallet_volume > 0 else 0

    return base_util, volume_util


def create_3d_plot(pallet, base_plan):
    plan = {
        "pallet_length": pallet["pallet_length"],
        "pallet_width": pallet["pallet_width"],
        "pallet_height": pallet["pallet_height"],
        "max_load_height": base_plan["max_load_height"],
        "overhang": base_plan["overhang"]
    }

    fig = go.Figure()
    color_map = {}
    legend = []

    px = 0
    py = 0
    pz = 0
    pl = plan["pallet_length"]
    pw = plan["pallet_width"]
    ph = plan["pallet_height"]

    fig.add_trace(go.Mesh3d(
        x=[px, px+pl, px+pl, px, px, px+pl, px+pl, px],
        y=[py, py, py+pw, py+pw, py, py, py+pw, py+pw],
        z=[pz, pz, pz, pz, pz+ph, pz+ph, pz+ph, pz+ph],
        i=[0, 0, 0, 1, 4, 4, 5, 2, 6, 3, 7, 1],
        j=[1, 2, 3, 2, 5, 6, 6, 3, 7, 0, 4, 5],
        k=[2, 3, 1, 0, 6, 7, 1, 0, 3, 4, 5, 6],
        opacity=0.45,
        color="#9ca3af",
        flatshading=True,
        name="Pallet",
        showscale=False
    ))

    def add_box_edges(x, y, z, dx, dy, dz, color):
        corners = [
            (x, y, z),
            (x + dx, y, z),
            (x + dx, y + dy, z),
            (x, y + dy, z),
            (x, y, z + dz),
            (x + dx, y, z + dz),
            (x + dx, y + dy, z + dz),
            (x, y + dy, z + dz),
        ]

        edges = [
            (0, 1), (1, 2), (2, 3), (3, 0),
            (4, 5), (5, 6), (6, 7), (7, 4),
            (0, 4), (1, 5), (2, 6), (3, 7)
        ]

        for a, b in edges:
            fig.add_trace(go.Scatter3d(
                x=[corners[a][0], corners[b][0]],
                y=[corners[a][1], corners[b][1]],
                z=[corners[a][2], corners[b][2]],
                mode="lines",
                line=dict(color=color, width=7),
                showlegend=False,
                hoverinfo="skip"
            ))

    for box in pallet["boxes"]:
        if box["box_name"] not in color_map:
            color_map[box["box_name"]] = BOX_COLORS[len(color_map) % len(BOX_COLORS)]
            legend.append({
                "name": box["box_name"],
                "color": color_map[box["box_name"]]
            })

        color = color_map[box["box_name"]]

        x = box["x"]
        y = box["y"]
        z = box["z"]
        dx = box["placed_length"]
        dy = box["placed_width"]
        dz = box["height"]

        fig.add_trace(go.Mesh3d(
            x=[x, x+dx, x+dx, x],
            y=[y, y, y+dy, y+dy],
            z=[z+dz, z+dz, z+dz, z+dz],
            i=[0, 0],
            j=[1, 2],
            k=[2, 3],
            opacity=1.0,
            color=color,
            flatshading=True,
            name=box["box_name"],
            showscale=False
        ))

        add_box_edges(x, y, z, dx, dy, dz, color)

    max_dim = max(plan["pallet_length"], plan["pallet_width"], plan["max_load_height"])

    fig.update_layout(
        scene=dict(
            xaxis_title="Length",
            yaxis_title="Width",
            zaxis_title="Height",
            xaxis=dict(range=[0, plan["pallet_length"]], showgrid=True, zeroline=False),
            yaxis=dict(range=[0, plan["pallet_width"]], showgrid=True, zeroline=False),
            zaxis=dict(range=[0, plan["max_load_height"]], showgrid=True, zeroline=False),
            aspectmode="manual",
            aspectratio=dict(
                x=plan["pallet_length"] / max_dim,
                y=plan["pallet_width"] / max_dim,
                z=plan["max_load_height"] / max_dim
            ),
            camera=dict(
                eye=dict(x=1.7, y=1.5, z=1.1)
            ),
            bgcolor="rgba(0,0,0,0)"
        ),
        margin=dict(l=0, r=0, t=20, b=0),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color="white"),
        showlegend=False
    )

    return pio.to_html(fig, full_html=False), legend, []


def build_technical_view_data(pallet):
    grouped = {}
    for box in pallet["boxes"]:
        name = box["box_name"]
        grouped[name] = grouped.get(name, 0) + 1

    return {
        "pallet_type": pallet["pallet_type"],
        "pallet_length": pallet["pallet_length"],
        "pallet_width": pallet["pallet_width"],
        "pallet_height": pallet["pallet_height"],
        "weight": round(pallet["weight"], 2),
        "used_height": pallet["used_height"],
        "boxes": pallet["boxes"],
        "grouped_boxes": grouped
    }


@app.route("/")
def dashboard():
    conn = get_conn()
    plans = conn.execute("SELECT * FROM plans ORDER BY id DESC").fetchall()
    conn.close()
    return render_template("dashboard.html", plans=plans)


@app.route("/box-library", methods=["GET", "POST"])
def box_library():
    conn = get_conn()

    if request.method == "POST":
        conn.execute("""
            INSERT INTO box_library (name, length, width, height, weight)
            VALUES (?, ?, ?, ?, ?)
        """, (
            request.form["name"],
            int(request.form["length"]),
            int(request.form["width"]),
            int(request.form["height"]),
            float(request.form["weight"])
        ))
        conn.commit()
        return redirect(url_for("box_library"))

    boxes = conn.execute("SELECT * FROM box_library ORDER BY name").fetchall()
    conn.close()
    return render_template("box_library.html", boxes=boxes)


@app.route("/pallet-library", methods=["GET", "POST"])
def pallet_library():
    conn = get_conn()

    if request.method == "POST":
        conn.execute("""
            INSERT INTO pallet_library (name, length, width, height, max_weight)
            VALUES (?, ?, ?, ?, ?)
        """, (
            request.form["name"],
            int(request.form["length"]),
            int(request.form["width"]),
            int(request.form["height"]),
            float(request.form["max_weight"])
        ))
        conn.commit()
        return redirect(url_for("pallet_library"))

    pallets = conn.execute("SELECT * FROM pallet_library ORDER BY name").fetchall()
    conn.close()
    return render_template("pallet_library.html", pallets=pallets)


@app.route("/delete-box/<int:box_id>", methods=["POST"])
def delete_box(box_id):
    conn = get_conn()
    conn.execute("DELETE FROM box_library WHERE id = ?", (box_id,))
    conn.commit()
    conn.close()
    return redirect(url_for("box_library"))


@app.route("/delete-pallet-library/<int:pallet_id>", methods=["POST"])
def delete_pallet_library(pallet_id):
    conn = get_conn()
    conn.execute("DELETE FROM pallet_library WHERE id = ?", (pallet_id,))
    conn.commit()
    conn.close()
    return redirect(url_for("pallet_library"))


@app.route("/new-plan", methods=["GET", "POST"])
def new_plan():
    conn = get_conn()
    library_boxes = conn.execute("SELECT * FROM box_library ORDER BY name").fetchall()
    pallet_library_rows = conn.execute("SELECT * FROM pallet_library ORDER BY name").fetchall()

    if request.method == "POST":
        c = conn.cursor()
        c.execute("""
            INSERT INTO plans (
                plan_name, loading_target, max_load_height, overhang, allow_rotation, created_at
            ) VALUES (?, ?, ?, ?, ?, ?)
        """, (
            request.form["plan_name"],
            request.form["loading_target"],
            int(request.form["max_load_height"]),
            int(request.form["overhang"]),
            1 if request.form.get("allow_rotation") == "yes" else 0,
            datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ))
        plan_id = c.lastrowid

        for pallet_id in request.form.getlist("allowed_pallets"):
            conn.execute("""
                INSERT INTO plan_allowed_pallets (plan_id, pallet_library_id)
                VALUES (?, ?)
            """, (plan_id, int(pallet_id)))

        names = request.form.getlist("box_name")
        lengths = request.form.getlist("box_length")
        widths = request.form.getlist("box_width")
        heights = request.form.getlist("box_height")
        weights = request.form.getlist("box_weight")
        qtys = request.form.getlist("box_qty")

        for i in range(len(names)):
            if names[i].strip():
                conn.execute("""
                    INSERT INTO plan_boxes (
                        plan_id, box_name, length, width, height, weight, qty
                    ) VALUES (?, ?, ?, ?, ?, ?, ?)
                """, (
                    plan_id,
                    names[i].strip(),
                    int(lengths[i]),
                    int(widths[i]),
                    int(heights[i]),
                    float(weights[i]),
                    int(qtys[i])
                ))

        conn.commit()
        conn.close()
        return redirect(url_for("plan_detail", plan_id=plan_id))

    conn.close()
    return render_template(
        "new_plan.html",
        library_boxes=library_boxes,
        pallet_library=pallet_library_rows,
        loading_targets=LOADING_TARGETS
    )


@app.route("/plan/<int:plan_id>")
def plan_detail(plan_id):
    conn = get_conn()
    plan = conn.execute("SELECT * FROM plans WHERE id = ?", (plan_id,)).fetchone()

    if not plan:
        conn.close()
        return redirect(url_for("dashboard"))

    boxes = conn.execute("SELECT * FROM plan_boxes WHERE plan_id = ?", (plan_id,)).fetchall()
    allowed_pallets = conn.execute("""
        SELECT pallet_library.*
        FROM plan_allowed_pallets
        JOIN pallet_library ON plan_allowed_pallets.pallet_library_id = pallet_library.id
        WHERE plan_allowed_pallets.plan_id = ?
        ORDER BY pallet_library.name
    """, (plan_id,)).fetchall()

    total_boxes = sum(box["qty"] for box in boxes)
    total_weight = sum(box["qty"] * box["weight"] for box in boxes)

    conn.close()

    return render_template(
        "plan_detail.html",
        plan=plan,
        boxes=boxes,
        allowed_pallets=allowed_pallets,
        total_boxes=total_boxes,
        total_weight=round(total_weight, 2)
    )


@app.route("/calculate-plan/<int:plan_id>")
def calculate_plan(plan_id):
    conn = get_conn()

    base_plan = conn.execute("SELECT * FROM plans WHERE id = ?", (plan_id,)).fetchone()
    if not base_plan:
        conn.close()
        return redirect(url_for("dashboard"))

    boxes = conn.execute("SELECT * FROM plan_boxes WHERE plan_id = ?", (plan_id,)).fetchall()
    allowed_pallets = conn.execute("""
        SELECT pallet_library.*
        FROM plan_allowed_pallets
        JOIN pallet_library ON plan_allowed_pallets.pallet_library_id = pallet_library.id
        WHERE plan_allowed_pallets.plan_id = ?
        ORDER BY pallet_library.name
    """, (plan_id,)).fetchall()

    conn.close()

    unique_box_types = {box["box_name"] for box in boxes}
    single_box_debug = None
    all_pallet_debug = []
    pallets = []
    selected_plan = base_plan

    if len(unique_box_types) == 1:
        merged_box = merge_same_box_type_rows(boxes)

        for pallet_row in allowed_pallets:
            cap = single_box_capacity(base_plan, pallet_row, merged_box)
            if cap:
                all_pallet_debug.append({
                    "pallet_name": pallet_row["name"],
                    "status": "fit",
                    **cap
                })
            else:
                all_pallet_debug.append({
                    "pallet_name": pallet_row["name"],
                    "status": "no fit"
                })

        strict_result = solve_single_box_type_strict(base_plan, allowed_pallets, merged_box)

        if strict_result:
            selected_plan = strict_result["plan"]
            pallets = strict_result["pallets"]
            single_box_debug = strict_result["debug"]

    else:
        pallets = solve_multi_box_simple(base_plan, allowed_pallets, boxes)

    pallet_summaries = []
    editable_summary = []
    serialized_pallets = []

    for idx, pallet in enumerate(pallets, start=1):
        grouped = {}
        for box in pallet["boxes"]:
            grouped[box["box_name"]] = grouped.get(box["box_name"], 0) + 1

        editable_summary.append({
            "index": idx,
            "grouped_boxes": grouped,
            "pallet_type": pallet["pallet_type"]
        })

        base_util, volume_util = base_and_volume_utilization(pallet, selected_plan)

        pallet_summaries.append({
            "index": idx,
            "pallet_type": pallet["pallet_type"],
            "weight": round(pallet["weight"], 2),
            "used_height": pallet["used_height"],
            "box_count": len(pallet["boxes"]),
            "grouped_boxes": grouped,
            "base_layer_utilization": base_util,
            "volume_utilization": volume_util,
            "overlap_count": 0
        })

        serialized_pallets.append(copy.deepcopy(pallet))

    session["editable_layout"] = editable_summary
    session["last_calculated_pallets"] = serialized_pallets
    session["last_selected_plan"] = dict(selected_plan)

    return render_template(
        "calculation_result.html",
        plan=selected_plan,
        plan_id=plan_id,
        total_pallets=len(pallets),
        pallet_summaries=pallet_summaries,
        single_box_debug=single_box_debug,
        all_pallet_debug=all_pallet_debug
    )


@app.route("/pallet-3d/<int:plan_id>/<int:pallet_index>")
def pallet_3d(plan_id, pallet_index):
    pallets = session.get("last_calculated_pallets", [])
    selected_plan = session.get("last_selected_plan")

    if not selected_plan or pallet_index < 1 or pallet_index > len(pallets):
        return redirect(url_for("plan_detail", plan_id=plan_id))

    pallet = pallets[pallet_index - 1]
    plot_html, legend, overlaps = create_3d_plot(pallet, selected_plan)

    return render_template(
        "pallet_3d.html",
        plan=selected_plan,
        plan_id=plan_id,
        pallet_index=pallet_index,
        pallet=pallet,
        plot_html=plot_html,
        legend=legend,
        overlaps=overlaps
    )


@app.route("/pallet-technical/<int:plan_id>/<int:pallet_index>")
def pallet_technical(plan_id, pallet_index):
    pallets = session.get("last_calculated_pallets", [])
    selected_plan = session.get("last_selected_plan")

    if not selected_plan or pallet_index < 1 or pallet_index > len(pallets):
        return redirect(url_for("plan_detail", plan_id=plan_id))

    pallet = pallets[pallet_index - 1]
    technical = build_technical_view_data(pallet)

    return render_template(
        "pallet_technical.html",
        plan=selected_plan,
        plan_id=plan_id,
        pallet_index=pallet_index,
        pallet=technical
    )


@app.route("/update-layout/<int:plan_id>", methods=["POST"])
def update_layout(plan_id):
    editable_layout = session.get("editable_layout", [])

    for pallet in editable_layout:
        for box_name in list(pallet["grouped_boxes"].keys()):
            field_name = f"pallet_{pallet['index']}_{box_name}"
            if field_name in request.form:
                pallet["grouped_boxes"][box_name] = int(request.form[field_name])

    session["editable_layout"] = editable_layout
    return redirect(url_for("show_edited_layout", plan_id=plan_id))


@app.route("/show-edited-layout/<int:plan_id>")
def show_edited_layout(plan_id):
    conn = get_conn()
    plan = conn.execute("SELECT * FROM plans WHERE id = ?", (plan_id,)).fetchone()

    if not plan:
        conn.close()
        return redirect(url_for("dashboard"))

    conn.close()
    editable_layout = session.get("editable_layout", [])

    return render_template(
        "edited_layout.html",
        plan=plan,
        editable_layout=editable_layout
    )


@app.route("/delete-plan/<int:plan_id>", methods=["POST"])
def delete_plan(plan_id):
    conn = get_conn()
    conn.execute("DELETE FROM plan_boxes WHERE plan_id = ?", (plan_id,))
    conn.execute("DELETE FROM plan_allowed_pallets WHERE plan_id = ?", (plan_id,))
    conn.execute("DELETE FROM plans WHERE id = ?", (plan_id,))
    conn.commit()
    conn.close()
    return redirect(url_for("dashboard"))


@app.route("/export-plan/<int:plan_id>")
def export_plan(plan_id):
    conn = get_conn()

    plan = conn.execute("SELECT * FROM plans WHERE id = ?", (plan_id,)).fetchone()
    if not plan:
        conn.close()
        return redirect(url_for("dashboard"))

    boxes = conn.execute("SELECT * FROM plan_boxes WHERE plan_id = ?", (plan_id,)).fetchall()
    allowed_pallets = conn.execute("""
        SELECT pallet_library.name
        FROM plan_allowed_pallets
        JOIN pallet_library ON plan_allowed_pallets.pallet_library_id = pallet_library.id
        WHERE plan_allowed_pallets.plan_id = ?
        ORDER BY pallet_library.name
    """, (plan_id,)).fetchall()

    conn.close()

    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Plan Summary"

    ws1.append(["Plan name", plan["plan_name"]])
    ws1.append(["Loading target", plan["loading_target"]])
    ws1.append(["Max load height", plan["max_load_height"]])
    ws1.append(["Allowed overhang", plan["overhang"]])
    ws1.append(["Allow rotation", "Yes" if plan["allow_rotation"] else "No"])
    ws1.append(["Created at", plan["created_at"]])

    ws1.append([])
    ws1.append(["Allowed pallet types"])
    for pallet in allowed_pallets:
        ws1.append([pallet["name"]])

    ws2 = wb.create_sheet("Boxes")
    ws2.append(["Box name", "Length", "Width", "Height", "Weight", "Qty"])

    for box in boxes:
        ws2.append([
            box["box_name"],
            box["length"],
            box["width"],
            box["height"],
            box["weight"],
            box["qty"]
        ])

    filename = f'plan_{plan_id}.xlsx'
    wb.save(filename)

    return send_file(filename, as_attachment=True)


init_db()

if __name__ == "__main__":
    app.run(debug=True)
