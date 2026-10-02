import math
import random
import time
import numpy as np
from pynput.mouse import Controller as MouseController, Button

mouse = MouseController()

def _bezier_point(p0, p1, p2, p3, t):
    """Calculates a point along a Cubic Bézier curve at step t (0.0 <= t <= 1.0)."""
    return (
        (1 - t) ** 3 * p0
        + 3 * (1 - t) ** 2 * t * p1
        + 3 * (1 - t) * t ** 2 * p2
        + t ** 3 * p3
    )

def _fitts_velocity_profile(t):
    """Minimum-jerk approximation: slow start, rapid transit, slow arrival."""
    return t * t * (3 - 2 * t)

def human_move(dest_x, dest_y, base_duration=0.55, allow_overshoot=True):
    """
    Moves mouse from current position to (dest_x, dest_y) mimicking
    biological motor control, variable acceleration, and overshoot.
    """
    start_x, start_y = mouse.position
    dx = dest_x - start_x
    dy = dest_y - start_y
    distance = math.hypot(dx, dy)

    if distance < 3:
        return

    # Fitts's law scaling
    duration = max(0.28, base_duration * (1.0 + math.log10(max(1.0, distance / 140.0))))

    # Overshoot probability on longer sweeps
    overshoot_point = None
    if allow_overshoot and distance > 100 and random.random() < 0.30:
        overshoot_mag = random.uniform(5.0, 14.0)
        angle = math.atan2(dy, dx) + random.uniform(-0.15, 0.15)
        overshoot_point = (
            dest_x + overshoot_mag * math.cos(angle),
            dest_y + overshoot_mag * math.sin(angle),
        )

    actual_target_x = overshoot_point[0] if overshoot_point else dest_x
    actual_target_y = overshoot_point[1] if overshoot_point else dest_y

    p0 = np.array([start_x, start_y], dtype=float)
    p3 = np.array([actual_target_x, actual_target_y], dtype=float)

    # Lateral control points to generate natural arm curvature
    perp_x = -dy / (distance + 1e-5)
    perp_y = dx / (distance + 1e-5)
    arc_intensity = random.uniform(-0.18, 0.18) * distance

    p1 = p0 + np.array([dx * random.uniform(0.2, 0.4), dy * random.uniform(0.2, 0.4)])
    p1 += np.array([perp_x * arc_intensity, perp_y * arc_intensity])

    p2 = p0 + np.array([dx * random.uniform(0.6, 0.8), dy * random.uniform(0.6, 0.8)])
    p2 += np.array([perp_x * (arc_intensity * 0.5), perp_y * (arc_intensity * 0.5)])

    steps = max(25, int(duration * 120))
    start_time = time.time()

    for step in range(steps + 1):
        raw_t = step / steps
        eased_t = _fitts_velocity_profile(raw_t)

        pos = _bezier_point(p0, p1, p2, p3, eased_t)

        # Micro-tremors (1px fine muscle noise)
        jitter_x = random.choice([-0.4, 0.0, 0.4]) if 0.15 < raw_t < 0.85 else 0.0
        jitter_y = random.choice([-0.4, 0.0, 0.4]) if 0.15 < raw_t < 0.85 else 0.0

        mouse.position = (int(round(pos[0] + jitter_x)), int(round(pos[1] + jitter_y)))

        target_time = start_time + (raw_t * duration)
        sleep_left = target_time - time.time()
        if sleep_left > 0:
            time.sleep(sleep_left)

    mouse.position = (int(round(actual_target_x)), int(round(actual_target_y)))

    # Secondary corrective homing movement if overshoot occurred
    if overshoot_point:
        time.sleep(random.uniform(0.05, 0.12))
        human_move(dest_x, dest_y, base_duration=0.20, allow_overshoot=False)

def human_click(target_x=None, target_y=None, button=Button.left, add_jitter=True):
    """Moves to coordinates and clicks with authentic physical hold times."""
    if target_x is not None and target_y is not None:
        if add_jitter:
            target_x += random.randint(-4, 4)
            target_y += random.randint(-3, 3)
        human_move(target_x, target_y)

    time.sleep(random.uniform(0.06, 0.14))

    # Real human click press-and-release dwell time (50ms - 120ms)
    mouse.press(button)
    hold_time = max(0.045, random.lognormvariate(math.log(0.075), 0.22))
    time.sleep(hold_time)
    mouse.release(button)

    time.sleep(random.uniform(0.08, 0.18))

def human_scroll(clicks, direction='down'):
    """Performs natural stepped scroll wheel movements with micro-delays."""
    sign = -1 if direction == 'down' else 1
    total_steps = abs(clicks)
    for _ in range(total_steps):
        mouse.scroll(0, sign)
        time.sleep(random.uniform(0.06, 0.14))

def human_reading_wander(center_x, center_y, passes=2):
    """Simulates eye-tracking cursor drift while reading or checking code."""
    cx, cy = center_x, center_y
    for _ in range(passes):
        wander_x = cx + random.randint(-60, 100)
        wander_y = cy + random.randint(30, 90)
        human_move(wander_x, wander_y, base_duration=random.uniform(0.6, 1.1), allow_overshoot=False)
        time.sleep(random.uniform(0.5, 1.2))
        cy = wander_y

def human_drag(start_x, start_y, end_x, end_y, duration=0.65):
    """Moves to start coordinates, drags across text, and releases."""
    human_move(start_x, start_y, base_duration=0.45)
    time.sleep(random.uniform(0.08, 0.16))

    mouse.press(Button.left)
    time.sleep(random.uniform(0.06, 0.12))

    p0 = np.array([start_x, start_y], dtype=float)
    p3 = np.array([end_x, end_y], dtype=float)
    dx = end_x - start_x
    dy = end_y - start_y

    ctrl_drift = random.uniform(-4, 4)
    p1 = p0 + np.array([dx * 0.35, dy * 0.35 + ctrl_drift])
    p2 = p0 + np.array([dx * 0.70, dy * 0.70 + ctrl_drift])

    steps = max(20, int(duration * 100))
    start_time = time.time()

    for i in range(steps + 1):
        raw_t = i / steps
        eased_t = _fitts_velocity_profile(raw_t)
        pos = _bezier_point(p0, p1, p2, p3, eased_t)

        jitter_y = random.choice([-0.3, 0.0, 0.3])
        mouse.position = (int(round(pos[0])), int(round(pos[1] + jitter_y)))

        target_time = start_time + (raw_t * duration)
        sleep_dur = target_time - time.time()
        if sleep_dur > 0:
            time.sleep(sleep_dur)

    mouse.position = (int(end_x), int(end_y))
    time.sleep(random.uniform(0.08, 0.18))
    mouse.release(Button.left)
    time.sleep(random.uniform(0.12, 0.25))