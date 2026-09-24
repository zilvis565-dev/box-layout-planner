from flask import Flask, render_template, request, redirect, url_for, send_file, session
import sqlite3
from datetime import datetime
from openpyxl import Workbook
import plotly.graph_objects as go
import plotly.io as pio
import copy
import math

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
        existing = c.execute("""
            SELECT id FROM pallet_library WHERE name = ?
        """, (pallet["name"],)).fetchone()

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

    conn.commit()
    conn.close()


def expand_boxes(boxes):
    expanded = []
    for box in boxes:
        for _ in range(box["qty"]):
            expanded.append({
                "box_name": box["box_name"],
                "length": box["length"],
                "width": box["width"],
                "height": box["height"],
                "weight": box["weight"]
            })

    expanded.sort(
        key=lambda x: (
            x["length"] * x["width"],
            x["weight"],
            x["height"],
            x["length"] * x["width"] * x["height"]
        ),
        reverse=True
    )
    return expanded


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


def single_box_type_distribution(base_plan, pallet_row, box):
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

        if per_layer <= 0 or layers <= 0:
            continue

        by_height = per_layer * layers
        by_weight = int(plan["max_pallet_weight"] // box["weight"]) if box["weight"] > 0 else by_height
        per_pallet = min(by_height, by_weight)

        if per_pallet <= 0:
            continue

        pallets_needed = math.ceil(box["qty"] / per_pallet)

        candidate = {
            "orientation": (bl, bw),
            "per_row": int(per_row),
            "per_col": int(per_col),
            "per_layer": int(per_layer),
            "layers": int(layers),
            "per_pallet": int(per_pallet),
            "pallets_needed": int(pallets_needed)
        }

        if best is None or candidate["per_pallet"] > best["per_pallet"]:
            best = candidate

    if not best:
        return None

    pallets = []
    remaining = box["qty"]

    for _ in range(best["pallets_needed"]):
        pallet_qty = min(remaining, best["per_pallet"])
        remaining -= pallet_qty

        pallet = {
            "boxes": [],
            "weight": pallet_qty * box["weight"],
            "used_height": plan["pallet_height"],
            "layers": [],
            "pallet_type": plan["pallet_type"],
            "pallet_length": plan["pallet_length"],
            "pallet_width": plan["pallet_width"],
            "pallet_height": plan["pallet_height"],
            "max_weight": plan["max_pallet_weight"]
        }

        bl, bw = best["orientation"]
        qty_left = pallet_qty
        z = plan["pallet_height"]

        for _layer in range(best["layers"]):
            if qty_left <= 0:
                break

            layer_count = min(qty_left, best["per_layer"])
            qty_left -= layer_count

            x = 0
            y = 0
            count = 0

            for _r in range(best["per_col"]):
                x = 0
                for _c in range(best["per_row"]):
                    if count >= layer_count:
                        break

                    pallet["boxes"].append({
                        "box_name": box["box_name"],
                        "length": box["length"],
                        "width": box["width"],
                        "height": box["height"],
                        "weight": box["weight"],
                        "placed_length": bl,
                        "placed_width": bw,
                        "x": x,
                        "y": y,
                        "z": z
                    })

                    x += bl
                    count += 1

                y += bw
                if count >= layer_count:
                    break

            pallet["used_height"] = max(pallet["used_height"], z + box["height"])
            z += box["height"]

        pallets.append(pallet)

    return {
        "plan": plan,
        "pallets": pallets
    }


def try_place_box_in_pallet(box, pallet, plan):
    effective_length = plan["pallet_length"] + 2 * plan["overhang"]
    effective_width = plan["pallet_width"] + 2 * plan["overhang"]
    pallet_height = plan["pallet_height"]
    max_height = plan["max_load_height"]
    max_weight = plan["max_pallet_weight"]
    allow_rotation = bool(plan["allow_rotation"])

    if "boxes" not in pallet:
        pallet["boxes"] = []
    if "weight" not in pallet:
        pallet["weight"] = 0
    if "used_height" not in pallet:
        pallet["used_height"] = pallet_height
    if "layers" not in pallet:
        pallet["layers"] = []
    if "pallet_type" not in pallet:
        pallet["pallet_type"] = plan["pallet_type"]
    if "pallet_length" not in pallet:
        pallet["pallet_length"] = plan["pallet_length"]
    if "pallet_width" not in pallet:
        pallet["pallet_width"] = plan["pallet_width"]
    if "pallet_height" not in pallet:
        pallet["pallet_height"] = plan["pallet_height"]
    if "max_weight" not in pallet:
        pallet["max_weight"] = plan["max_pallet_weight"]

    orientations = [(box["length"], box["width"])]
    if allow_rotation and box["length"] != box["width"]:
        orientations.append((box["width"], box["length"]))

    orientations.sort(key=lambda o: o[0] * o[1], reverse=True)

    for idx, layer in enumerate(pallet["layers"]):
        free_spaces = sorted(layer["free_spaces"], key=lambda s: s[2] * s[3], reverse=True)
        layer_z = layer["z"]

        for sidx, space in enumerate(free_spaces):
            sx, sy, sw, sd = space

            for box_l, box_w in orientations:
                if (
                    box_l <= sw and
                    box_w <= sd and
                    layer_z + box["height"] <= max_height and
                    pallet["weight"] + box["weight"] <= max_weight
                ):
                    placed_box = {
                        **box,
                        "placed_length": box_l,
                        "placed_width": box_w,
                        "x": sx,
                        "y": sy,
                        "z": layer_z
                    }

                    new_spaces = free_spaces[:sidx] + free_spaces[sidx+1:]
                    right_space = (sx + box_l, sy, sw - box_l, box_w)
                    bottom_space = (sx, sy + box_w, sw, sd - box_w)

                    if right_space[2] > 0 and right_space[3] > 0:
                        new_spaces.append(right_space)
                    if bottom_space[2] > 0 and bottom_space[3] > 0:
                        new_spaces.append(bottom_space)

                    new_spaces.sort(key=lambda s: s[2] * s[3], reverse=True)
                    pallet["layers"][idx]["free_spaces"] = new_spaces

                    pallet["boxes"].append(placed_box)
                    pallet["weight"] += box["weight"]
                    pallet["used_height"] = max(pallet["used_height"], layer_z + box["height"])
                    return True

    current_top = pallet_height
    if pallet["layers"]:
        current_top = max(layer["z"] + layer["height"] for layer in pallet["layers"])

    for box_l, box_w in orientations:
        if (
            current_top + box["height"] <= max_height and
            pallet["weight"] + box["weight"] <= max_weight and
            box_l <= effective_length and
            box_w <= effective_width
        ):
            new_layer = {
                "z": current_top,
                "height": box["height"],
                "free_spaces": []
            }

            placed_box = {
                **box,
                "placed_length": box_l,
                "placed_width": box_w,
                "x": 0,
                "y": 0,
                "z": current_top
            }

            right_space = (box_l, 0, effective_length - box_l, box_w)
            bottom_space = (0, box_w, effective_length, effective_width - box_w)

            if right_space[2] > 0 and right_space[3] > 0:
                new_layer["free_spaces"].append(right_space)
            if bottom_space[2] > 0 and bottom_space[3] > 0:
                new_layer["free_spaces"].append(bottom_space)

            new_layer["free_spaces"].sort(key=lambda s: s[2] * s[3], reverse=True)

            pallet["layers"].append(new_layer)
            pallet["boxes"].append(placed_box)
            pallet["weight"] += box["weight"]
            pallet["used_height"] = max(pallet["used_height"], current_top + box["height"])
            return True

    return False


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


def pallet_score(pallet, plan):
    base_util, volume_util = base_and_volume_utilization(pallet, plan)
    height_ratio = pallet["used_height"] / plan["max_load_height"] if plan["max_load_height"] > 0 else 0
    weight_ratio = pallet["weight"] / plan["max_pallet_weight"] if plan["max_pallet_weight"] > 0 else 0
    return (base_util / 100) * 2 + (volume_util / 100) + height_ratio + weight_ratio


def choose_best_new_pallet_for_box(box, base_plan, allowed_pallets):
    best = None

    for pallet_row in allowed_pallets:
        pallet_plan = build_plan_from_pallet(base_plan, pallet_row)
        new_pallet = {
            "boxes": [],
            "weight": 0,
            "used_height": pallet_plan["pallet_height"],
            "layers": [],
            "pallet_type": pallet_plan["pallet_type"],
            "pallet_length": pallet_plan["pallet_length"],
            "pallet_width": pallet_plan["pallet_width"],
            "pallet_height": pallet_plan["pallet_height"],
            "max_weight": pallet_plan["max_pallet_weight"]
        }

        if try_place_box_in_pallet(box, new_pallet, pallet_plan):
            score = pallet_score(new_pallet, pallet_plan)
            candidate = {
                "pallet": new_pallet,
                "plan": pallet_plan,
                "score": score
            }

            if best is None or candidate["score"] > best["score"]:
                best = candidate

    return best


def calculate_mixed_pallet_distribution(base_plan, allowed_pallets, boxes):
    expanded_boxes = expand_boxes(boxes)
    pallets = []

    for box in expanded_boxes:
        best_existing_index = None
        best_existing_score = -1

        for idx, pallet in enumerate(pallets):
            pallet_plan = {
                "pallet_type": pallet["pallet_type"],
                "pallet_length": pallet["pallet_length"],
                "pallet_width": pallet["pallet_width"],
                "pallet_height": pallet["pallet_height"],
                "max_load_height": base_plan["max_load_height"],
                "max_pallet_weight": pallet["max_weight"],
                "overhang": base_plan["overhang"],
                "allow_rotation": base_plan["allow_rotation"],
                "loading_target": base_plan["loading_target"]
            }

            test_pallet = copy.deepcopy(pallet)
            if try_place_box_in_pallet(box, test_pallet, pallet_plan):
                score = pallet_score(test_pallet, pallet_plan)
                if score > best_existing_score:
                    best_existing_score = score
                    best_existing_index = idx

        best_new = choose_best_new_pallet_for_box(box, base_plan, allowed_pallets)

        if best_existing_index is not None and best_new is not None:
            if best_existing_score >= best_new["score"]:
                chosen_pallet = pallets[best_existing_index]
                chosen_plan = {
                    "pallet_type": chosen_pallet["pallet_type"],
                    "pallet_length": chosen_pallet["pallet_length"],
                    "pallet_width": chosen_pallet["pallet_width"],
                    "pallet_height": chosen_pallet["pallet_height"],
                    "max_load_height": base_plan["max_load_height"],
                    "max_pallet_weight": chosen_pallet["max_weight"],
                    "overhang": base_plan["overhang"],
                    "allow_rotation": base_plan["allow_rotation"],
                    "loading_target": base_plan["loading_target"]
                }
                try_place_box_in_pallet(box, chosen_pallet, chosen_plan)
            else:
                pallets.append(best_new["pallet"])
        elif best_existing_index is not None:
            chosen_pallet = pallets[best_existing_index]
            chosen_plan = {
                "pallet_type": chosen_pallet["pallet_type"],
                "pallet_length": chosen_pallet["pallet_length"],
                "pallet_width": chosen_pallet["pallet_width"],
                "pallet_height": chosen_pallet["pallet_height"],
                "max_load_height": base_plan["max_load_height"],
                "max_pallet_weight": chosen_pallet["max_weight"],
                "overhang": base_plan["overhang"],
                "allow_rotation": base_plan["allow_rotation"],
                "loading_target": base_plan["loading_target"]
            }
            try_place_box_in_pallet(box, chosen_pallet, chosen_plan)
        elif best_new is not None:
            pallets.append(best_new["pallet"])

    return pallets


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
            x=[x, x+dx, x+dx, x, x, x+dx, x+dx, x],
            y=[y, y, y+dy, y+dy, y, y, y+dy, y+dy],
            z=[z, z, z, z, z+dz, z+dz, z+dz, z+dz],
            i=[0, 0, 0, 1, 4, 4, 5, 2, 6, 3, 7, 1],
            j=[1, 2, 3, 2, 5, 6, 6, 3, 7, 0, 4, 5],
            k=[2, 3, 1, 0, 6, 7, 1, 0, 3, 4, 5, 6],
            opacity=0.9,
            color=color,
            flatshading=True,
            name=box["box_name"],
            showscale=False
        ))

    fig.update_layout(
        scene=dict(
            xaxis_title="Length",
            yaxis_title="Width",
            zaxis_title="Height",
            xaxis=dict(range=[0, plan["pallet_length"] + 2 * plan["overhang"]]),
            yaxis=dict(range=[0, plan["pallet_width"] + 2 * plan["overhang"]]),
            zaxis=dict(range=[0, plan["max_load_height"]]),
            aspectmode="data",
            bgcolor="rgba(0,0,0,0)"
        ),
        margin=dict(l=0, r=0, t=20, b=0),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color="white"),
        showlegend=False
    )

    return pio.to_html(fig, full_html=False), legend


@app.route("/")
def dashboard():
    conn = get_conn()
    plans = conn.execute("""
        SELECT * FROM plans ORDER BY id DESC
    """).fetchall()
    conn.close()
    return render_template("dashboard.html", plans=plans)


@app.route("/box-library", methods=["GET", "POST"])
def box_library():
    conn = get_conn()

    if request.method == "POST":
        name = request.form["name"]
        length = int(request.form["length"])
        width = int(request.form["width"])
        height = int(request.form["height"])
        weight = float(request.form["weight"])

        conn.execute("""
            INSERT INTO box_library (name, length, width, height, weight)
            VALUES (?, ?, ?, ?, ?)
        """, (name, length, width, height, weight))
        conn.commit()
        return redirect(url_for("box_library"))

    boxes = conn.execute("""
        SELECT * FROM box_library ORDER BY name
    """).fetchall()
    conn.close()

    return render_template("box_library.html", boxes=boxes)


@app.route("/pallet-library", methods=["GET", "POST"])
def pallet_library():
    conn = get_conn()

    if request.method == "POST":
        name = request.form["name"]
        length = int(request.form["length"])
        width = int(request.form["width"])
        height = int(request.form["height"])
        max_weight = float(request.form["max_weight"])

        conn.execute("""
            INSERT INTO pallet_library (name, length, width, height, max_weight)
            VALUES (?, ?, ?, ?, ?)
        """, (name, length, width, height, max_weight))
        conn.commit()
        return redirect(url_for("pallet_library"))

    pallets = conn.execute("""
        SELECT * FROM pallet_library ORDER BY name
    """).fetchall()
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
    library_boxes = conn.execute("""
        SELECT * FROM box_library ORDER BY name
    """).fetchall()

    pallet_library_rows = conn.execute("""
        SELECT * FROM pallet_library ORDER BY name
    """).fetchall()

    if request.method == "POST":
        plan_name = request.form["plan_name"]
        loading_target = request.form["loading_target"]
        max_load_height = int(request.form["max_load_height"])
        overhang = int(request.form["overhang"])
        allow_rotation = 1 if request.form.get("allow_rotation") == "yes" else 0

        c = conn.cursor()
        c.execute("""
            INSERT INTO plans (
                plan_name, loading_target, max_load_height, overhang, allow_rotation, created_at
            ) VALUES (?, ?, ?, ?, ?, ?)
        """, (
            plan_name,
            loading_target,
            max_load_height,
            overhang,
            allow_rotation,
            datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ))
        plan_id = c.lastrowid

        allowed_pallet_ids = request.form.getlist("allowed_pallets")
        for pallet_id in allowed_pallet_ids:
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

    plan = conn.execute("""
        SELECT * FROM plans WHERE id = ?
    """, (plan_id,)).fetchone()

    if not plan:
        conn.close()
        return redirect(url_for("dashboard"))

    boxes = conn.execute("""
        SELECT * FROM plan_boxes WHERE plan_id = ?
    """, (plan_id,)).fetchall()

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

    base_plan = conn.execute("""
        SELECT * FROM plans WHERE id = ?
    """, (plan_id,)).fetchone()

    if not base_plan:
        conn.close()
        return redirect(url_for("dashboard"))

    boxes = conn.execute("""
        SELECT * FROM plan_boxes WHERE plan_id = ?
    """, (plan_id,)).fetchall()

    allowed_pallets = conn.execute("""
        SELECT pallet_library.*
        FROM plan_allowed_pallets
        JOIN pallet_library ON plan_allowed_pallets.pallet_library_id = pallet_library.id
        WHERE plan_allowed_pallets.plan_id = ?
        ORDER BY pallet_library.name
    """, (plan_id,)).fetchall()

    conn.close()

    unique_box_types = {box["box_name"] for box in boxes}

    if len(unique_box_types) == 1 and len(boxes) == 1:
        best_single = None

        for pallet_row in allowed_pallets:
            result = single_box_type_distribution(base_plan, pallet_row, boxes[0])
            if not result:
                continue

            pallet_count = len(result["pallets"])
            if best_single is None or pallet_count < len(best_single["pallets"]):
                best_single = result
            elif best_single is not None and pallet_count == len(best_single["pallets"]):
                current_cap = len(result["pallets"][0]["boxes"]) if result["pallets"] else 0
                best_cap = len(best_single["pallets"][0]["boxes"]) if best_single["pallets"] else 0
                if current_cap > best_cap:
                    best_single = result

        if best_single:
            selected_plan = best_single["plan"]
            pallets = best_single["pallets"]
        else:
            pallets = []
            selected_plan = base_plan
    else:
        pallets = calculate_mixed_pallet_distribution(base_plan, allowed_pallets, boxes)
        selected_plan = base_plan

    pallet_summaries = []
    editable_summary = []

    for idx, pallet in enumerate(pallets, start=1):
        grouped = {}
        for box in pallet["boxes"]:
            name = box["box_name"]
            grouped[name] = grouped.get(name, 0) + 1

        editable_summary.append({
            "index": idx,
            "grouped_boxes": grouped,
            "pallet_type": pallet["pallet_type"]
        })

        plot_html, legend = create_3d_plot(pallet, selected_plan)
        base_util, volume_util = base_and_volume_utilization(pallet, selected_plan)

        pallet_summaries.append({
            "index": idx,
            "pallet_type": pallet["pallet_type"],
            "weight": round(pallet["weight"], 2),
            "used_height": pallet["used_height"],
            "box_count": len(pallet["boxes"]),
            "grouped_boxes": grouped,
            "plot_html": plot_html,
            "legend": legend,
            "base_layer_utilization": base_util,
            "volume_utilization": volume_util
        })

    session["editable_layout"] = editable_summary

    return render_template(
        "calculation_result.html",
        plan=selected_plan,
        plan_id=plan_id,
        total_pallets=len(pallets),
        pallet_summaries=pallet_summaries
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
    plan = conn.execute("""
        SELECT * FROM plans WHERE id = ?
    """, (plan_id,)).fetchone()

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

    plan = conn.execute("""
        SELECT * FROM plans WHERE id = ?
    """, (plan_id,)).fetchone()

    if not plan:
        conn.close()
        return redirect(url_for("dashboard"))

    boxes = conn.execute("""
        SELECT * FROM plan_boxes WHERE plan_id = ?
    """, (plan_id,)).fetchall()

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
