import random
import tkinter as tk


WINDOW_WIDTH = 420
WINDOW_HEIGHT = 700
ROAD_LEFT = 80
ROAD_RIGHT = 340
LANE_COUNT = 3
LANE_WIDTH = (ROAD_RIGHT - ROAD_LEFT) // LANE_COUNT
PLAYER_Y = WINDOW_HEIGHT - 120
ENEMY_SPAWN_INTERVAL_MS = 900
FRAME_INTERVAL_MS = 16


class CarRaceGame:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Car Race Game")
        self.root.resizable(False, False)

        self.score = 0
        self.best_score = 0
        self.level = 1
        self.speed = 6
        self.running = True

        self.enemies = []
        self.road_lines = []

        self.player_lane = 1
        self.player_car = None

        self.score_label = None
        self.level_label = None
        self.message_label = None
        self.canvas = None

        self.build_ui()
        self.draw_static_scene()
        self.create_player_car()

        self.root.bind("<Left>", self.move_left)
        self.root.bind("<Right>", self.move_right)
        self.root.bind("a", self.move_left)
        self.root.bind("d", self.move_right)
        self.root.bind("r", self.restart)

        self.schedule_enemy_spawn()
        self.game_loop()

    def build_ui(self):
        wrapper = tk.Frame(self.root, bg="#1f2937", padx=12, pady=12)
        wrapper.pack(fill="both", expand=True)

        title = tk.Label(
            wrapper,
            text="Car Race",
            font=("Arial", 22, "bold"),
            bg="#1f2937",
            fg="#f9fafb",
        )
        title.pack(pady=(0, 10))

        info = tk.Frame(wrapper, bg="#111827", padx=10, pady=8)
        info.pack(fill="x")

        self.score_label = tk.Label(
            info,
            text="Score: 0",
            font=("Arial", 12, "bold"),
            bg="#111827",
            fg="#93c5fd",
        )
        self.score_label.pack(side="left")

        self.level_label = tk.Label(
            info,
            text="Level: 1",
            font=("Arial", 12, "bold"),
            bg="#111827",
            fg="#86efac",
        )
        self.level_label.pack(side="right")

        self.canvas = tk.Canvas(
            wrapper,
            width=WINDOW_WIDTH,
            height=WINDOW_HEIGHT,
            bg="#0f172a",
            highlightthickness=0,
        )
        self.canvas.pack(pady=10)

        self.message_label = tk.Label(
            wrapper,
            text="Use Left/Right arrow keys or A/D to move",
            font=("Arial", 11),
            bg="#1f2937",
            fg="#d1d5db",
        )
        self.message_label.pack()

    def draw_static_scene(self):
        self.canvas.create_rectangle(0, 0, ROAD_LEFT, WINDOW_HEIGHT, fill="#14532d", outline="#14532d")
        self.canvas.create_rectangle(ROAD_RIGHT, 0, WINDOW_WIDTH, WINDOW_HEIGHT, fill="#14532d", outline="#14532d")
        self.canvas.create_rectangle(ROAD_LEFT, 0, ROAD_RIGHT, WINDOW_HEIGHT, fill="#374151", outline="#374151")

        for i in range(0, WINDOW_HEIGHT, 70):
            line = self.canvas.create_rectangle(
                (ROAD_LEFT + ROAD_RIGHT) // 2 - 4,
                i,
                (ROAD_LEFT + ROAD_RIGHT) // 2 + 4,
                i + 40,
                fill="#f9fafb",
                outline="#f9fafb",
            )
            self.road_lines.append(line)

        self.canvas.create_line(ROAD_LEFT, 0, ROAD_LEFT, WINDOW_HEIGHT, fill="#fbbf24", width=3)
        self.canvas.create_line(ROAD_RIGHT, 0, ROAD_RIGHT, WINDOW_HEIGHT, fill="#fbbf24", width=3)

    def lane_center_x(self, lane: int) -> int:
        return ROAD_LEFT + lane * LANE_WIDTH + LANE_WIDTH // 2

    def create_player_car(self):
        x = self.lane_center_x(self.player_lane)
        self.player_car = self.draw_car(x, PLAYER_Y, "#3b82f6", "#1d4ed8", "player")

    def draw_car(self, x: int, y: int, body_color: str, roof_color: str, tag: str):
        body = self.canvas.create_rectangle(x - 22, y - 40, x + 22, y + 40, fill=body_color, outline="#111827", width=2, tags=tag)
        roof = self.canvas.create_rectangle(x - 14, y - 18, x + 14, y + 20, fill=roof_color, outline=roof_color, tags=tag)
        wheel1 = self.canvas.create_rectangle(x - 26, y - 30, x - 22, y - 14, fill="#111827", outline="#111827", tags=tag)
        wheel2 = self.canvas.create_rectangle(x + 22, y - 30, x + 26, y - 14, fill="#111827", outline="#111827", tags=tag)
        wheel3 = self.canvas.create_rectangle(x - 26, y + 14, x - 22, y + 30, fill="#111827", outline="#111827", tags=tag)
        wheel4 = self.canvas.create_rectangle(x + 22, y + 14, x + 26, y + 30, fill="#111827", outline="#111827", tags=tag)
        return [body, roof, wheel1, wheel2, wheel3, wheel4]

    def move_car_to_lane(self, car_parts, lane: int, y: int):
        x = self.lane_center_x(lane)
        boxes = [
            (x - 22, y - 40, x + 22, y + 40),
            (x - 14, y - 18, x + 14, y + 20),
            (x - 26, y - 30, x - 22, y - 14),
            (x + 22, y - 30, x + 26, y - 14),
            (x - 26, y + 14, x - 22, y + 30),
            (x + 22, y + 14, x + 26, y + 30),
        ]
        for part, coords in zip(car_parts, boxes):
            self.canvas.coords(part, *coords)

    def move_left(self, _event=None):
        if not self.running:
            return
        if self.player_lane > 0:
            self.player_lane -= 1
            self.move_car_to_lane(self.player_car, self.player_lane, PLAYER_Y)

    def move_right(self, _event=None):
        if not self.running:
            return
        if self.player_lane < LANE_COUNT - 1:
            self.player_lane += 1
            self.move_car_to_lane(self.player_car, self.player_lane, PLAYER_Y)

    def spawn_enemy(self):
        if not self.running:
            return

        lane = random.randint(0, LANE_COUNT - 1)
        y = -80
        body = random.choice(["#ef4444", "#f97316", "#a855f7", "#22c55e"])
        roof = random.choice(["#7f1d1d", "#7c2d12", "#4c1d95", "#14532d"])
        car_parts = self.draw_car(self.lane_center_x(lane), y, body, roof, "enemy")
        self.enemies.append({"lane": lane, "y": y, "parts": car_parts})

    def schedule_enemy_spawn(self):
        self.spawn_enemy()
        interval = max(350, ENEMY_SPAWN_INTERVAL_MS - (self.level - 1) * 70)
        self.root.after(interval, self.schedule_enemy_spawn)

    def update_road_lines(self):
        for line in self.road_lines:
            self.canvas.move(line, 0, self.speed)
            x1, y1, x2, y2 = self.canvas.coords(line)
            if y1 > WINDOW_HEIGHT:
                self.canvas.coords(line, x1, -40, x2, 0)

    def update_enemies(self):
        to_remove = []
        for enemy in self.enemies:
            enemy["y"] += self.speed
            self.move_car_to_lane(enemy["parts"], enemy["lane"], int(enemy["y"]))

            if enemy["y"] - 40 > WINDOW_HEIGHT:
                to_remove.append(enemy)
                self.score += 1
                if self.score % 10 == 0:
                    self.level += 1
                    self.speed = min(15, self.speed + 1)

        for enemy in to_remove:
            for part in enemy["parts"]:
                self.canvas.delete(part)
            self.enemies.remove(enemy)

    def player_bbox(self):
        body = self.player_car[0]
        return self.canvas.coords(body)

    def check_collision(self):
        p_left, p_top, p_right, p_bottom = self.player_bbox()

        for enemy in self.enemies:
            e_left, e_top, e_right, e_bottom = self.canvas.coords(enemy["parts"][0])
            overlap = not (p_right < e_left or p_left > e_right or p_bottom < e_top or p_top > e_bottom)
            if overlap:
                return True
        return False

    def update_hud(self):
        self.score_label.config(text=f"Score: {self.score}  Best: {self.best_score}")
        self.level_label.config(text=f"Level: {self.level}  Speed: {self.speed}")

    def game_over(self):
        self.running = False
        self.best_score = max(self.best_score, self.score)
        self.update_hud()
        self.message_label.config(text="Crash! Press R to restart", fg="#fecaca")
        self.canvas.create_text(
            WINDOW_WIDTH // 2,
            WINDOW_HEIGHT // 2,
            text="GAME OVER",
            font=("Arial", 36, "bold"),
            fill="#fee2e2",
            tags="game_over",
        )

    def restart(self, _event=None):
        if self.running:
            return

        for enemy in self.enemies:
            for part in enemy["parts"]:
                self.canvas.delete(part)
        self.enemies.clear()

        self.canvas.delete("game_over")
        self.score = 0
        self.level = 1
        self.speed = 6
        self.player_lane = 1
        self.move_car_to_lane(self.player_car, self.player_lane, PLAYER_Y)
        self.running = True
        self.message_label.config(text="Use Left/Right arrow keys or A/D to move", fg="#d1d5db")
        self.update_hud()

    def game_loop(self):
        self.update_road_lines()

        if self.running:
            self.update_enemies()
            if self.check_collision():
                self.game_over()
            self.update_hud()

        self.root.after(FRAME_INTERVAL_MS, self.game_loop)


def main():
    root = tk.Tk()
    CarRaceGame(root)
    root.mainloop()


if __name__ == "__main__":
    main()
