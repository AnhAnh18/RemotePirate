from pynput.mouse import Listener
import pygetwindow as gw
import pyautogui
import keyboard
import threading
import multiprocessing
import re
import os
from pywinauto import Desktop, Application


def get_chrome_windows_with_profile():
    """
    Lấy danh sách các cửa sổ Chrome kèm thông tin profile và HWND.
    Xử lý trường hợp không lấy được profile name (Unknown).
    """
    chrome_windows = []
    windows = gw.getAllWindows()

    profile_counts = {}  # Theo dõi số lần xuất hiện của mỗi profile

    for win in windows:
        if "Chrome" in win.title and " - Google Chrome" in win.title:
            hwnd = win._hWnd
            try:
                # Lấy đường dẫn của command line
                app = Application(backend="uia").connect(handle=hwnd)
                command_line = app.window(handle=hwnd).get_properties()['cached_command_line']

                # Trích xuất tên profile từ command line
                match = re.search(r"--profile-directory=[\"]?([^\" ]+)", command_line)
                if match:
                    profile_name = match.group(1)
                else:
                    profile_name = "Unknown"  # Gán nhãn "Unknown" nếu không lấy được

                # Xử lý trùng lặp profile
                if profile_name not in profile_counts:
                    profile_counts[profile_name] = 0
                profile_counts[profile_name] += 1

                # Thêm số thứ tự vào tên profile nếu trùng lặp
                if profile_counts[profile_name] > 1:
                    profile_name_with_count = f"{profile_name} ({profile_counts[profile_name]})"
                else:
                    profile_name_with_count = profile_name

                chrome_windows.append((win.title, hwnd, profile_name_with_count))

            except Exception as e:
                print(f"Lỗi khi lấy thông tin profile cho cửa sổ có HWND {hwnd}: {e}")
                chrome_windows.append((win.title, hwnd, "Unknown"))

    return chrome_windows


# Hàm lắng nghe sự kiện click chuột (chạy trong một process riêng)
def mouse_listener_process(main_window_hwnd, other_windows_hwnd, queue):
    def on_click(x, y, button, pressed):
        if pressed:
            try:
                # Kiểm tra xem sự kiện click có xảy ra trên cửa sổ chính hay không
                active_window = gw.getActiveWindow()
                if active_window and active_window._hWnd == main_window_hwnd:
                    print(f"Chuẩn bị lấy toạ độ chuột: ({x}, {y}) trên cửa sổ chính có HWND {main_window_hwnd}")

                    # Gửi dữ liệu click chuột về luồng chính thông qua Queue
                    # Bao gồm: tọa độ click, danh sách HWND cửa sổ phụ
                    queue.put((x, y, other_windows_hwnd))
                else:
                    print("Click chuột không xảy ra trên cửa sổ chính, bỏ qua.")

            except Exception as e:
                print(f"Lỗi khi xử lý sự kiện click chuột: {e}")

    with Listener(on_click=on_click) as listener:
        listener.join()

# Hàm xử lý click chuột (chạy trong luồng chính)
def process_mouse_clicks(queue):
    while True:
        try:
            # Lấy dữ liệu click chuột từ Queue
            x, y, other_windows_hwnd = queue.get()

            # Đồng bộ click trên các cửa sổ phụ sử dụng pywinauto
            app = Application(backend="uia")
            for hwnd in other_windows_hwnd:
                try:
                    # Kết nối tới cửa sổ Chrome phụ
                    app.connect(handle=hwnd, found_index=0)
                    # Lấy cửa sổ
                    window = app.window(handle=hwnd)
                    # Đảm bảo cửa sổ ở trạng thái bình thường (không minimized)
                    window.restore()

                    # Thực hiện click mà KHÔNG cần kích hoạt
                    window.click_input(coords=(x, y))
                    print(f"Đã click vào cửa sổ phụ có HWND '{hwnd}' tại vị trí ({x}, {y})")

                except Exception as e:
                    print(f"Lỗi khi click vào cửa sổ phụ có HWND '{hwnd}': {e}")

        except Exception as e:
            print(f"Lỗi khi xử lý click chuột: {e}")

# Hàm lắng nghe và đồng bộ phím (sử dụng window.type_keys() để không cần focus)
def listen_keyboard(main_window_hwnd, other_windows_hwnd):
    while True:
        event = keyboard.read_event()
        if event.event_type == "down":
            key = event.name
            try:
                # Kiểm tra xem cửa sổ chính có đang active không
                active_window = gw.getActiveWindow()
                if active_window and active_window._hWnd == main_window_hwnd:
                    print(f"Vừa nhấn nút: {key} trên cửa sổ chính")

                    # Gửi phím tới các cửa sổ phụ (không kích hoạt)
                    app = Application(backend="uia")
                    for hwnd in other_windows_hwnd:
                        try:
                            # Kết nối tới cửa sổ phụ
                            app.connect(handle=hwnd)
                            window = app.window(handle=hwnd)

                            # Gửi phím mà không cần kích hoạt
                            window.type_keys(key)

                            print(f"Đã đồng bộ phím '{key}' tới cửa sổ phụ có HWND '{hwnd}'")
                        except Exception as e:
                            print(f"Lỗi khi gửi phím tới cửa sổ có HWND '{hwnd}': {e}")
                else:
                    print("Cửa sổ chính không active, bỏ qua đồng bộ phím.")
            except Exception as e:
                print(f"Lỗi khi đồng bộ phím: {e}")

# Chương trình chính
def main():
    # Lấy danh sách các cửa sổ Chrome kèm HWND và profile
    chrome_windows_with_profile_hwnd = get_chrome_windows_with_profile()
    if not chrome_windows_with_profile_hwnd:
        print("Không tìm thấy cửa sổ Chrome nào đang mở!")
        return

    # Tạo danh sách profile duy nhất kèm số thứ tự
    profiles = []
    profile_counts = {}
    for _, _, profile in chrome_windows_with_profile_hwnd:
        if profile not in profile_counts:
            profile_counts[profile] = 0
            profiles.append(profile)
        profile_counts[profile] += 1

    print("Danh sách profiles:")
    for i, profile in enumerate(profiles):
        print(f"{i + 1}. {profile}")

    # Yêu cầu người dùng chọn profile từ danh sách
    while True:
        try:
            selected_profile_index = int(input("Nhập số thứ tự của profile cho cửa sổ chính: ")) - 1
            if 0 <= selected_profile_index < len(profiles):
                selected_profile = profiles[selected_profile_index]
                break
            else:
                print("Số thứ tự không hợp lệ. Vui lòng nhập lại.")
        except ValueError:
            print("Vui lòng nhập một số nguyên.")

    # Lọc danh sách cửa sổ theo profile đã chọn
    windows_with_selected_profile = [
        (title, hwnd)
        for title, hwnd, profile in chrome_windows_with_profile_hwnd
        if profile == selected_profile
    ]

    # Hiển thị danh sách cửa sổ Chrome cùng profile để người dùng chọn cửa sổ chính
    print(f"\nCửa sổ Chrome với profile '{selected_profile}':")
    for i, (title, hwnd) in enumerate(windows_with_selected_profile):
        print(f"{i + 1}. {title} (HWND: {hwnd})")

    # Yêu cầu người dùng chọn cửa sổ chính từ danh sách theo số thứ tự
    while True:
        try:
            main_window_index = int(input("Nhập số thứ tự của cửa sổ chính: ")) - 1
            if 0 <= main_window_index < len(windows_with_selected_profile):
                break
            else:
                print("Số thứ tự không hợp lệ. Vui lòng nhập lại.")
        except ValueError:
            print("Vui lòng nhập một số nguyên.")

    main_window_title, main_window_hwnd = windows_with_selected_profile[main_window_index]

    # Xác định các cửa sổ phụ (loại bỏ cửa sổ chính)
    other_windows_hwnd = [
        hwnd
        for title, hwnd, profile in chrome_windows_with_profile_hwnd
        if hwnd != main_window_hwnd
    ]

    print(f"\nCửa sổ chính: {main_window_title} (HWND: {main_window_hwnd}) - Profile: {selected_profile}")
    print(f"Các cửa sổ phụ (HWND): {other_windows_hwnd}")

    # Tạo Queue để truyền dữ liệu click chuột
    queue = multiprocessing.Queue()

    # Tạo process lắng nghe chuột
    mouse_process = multiprocessing.Process(target=mouse_listener_process,
                                            args=(main_window_hwnd, other_windows_hwnd, queue))

    # Tạo thread xử lý click chuột (chạy trong luồng chính)
    mouse_click_thread = threading.Thread(target=process_mouse_clicks, args=(queue,))

    # Tạo thread lắng nghe bàn phím
    keyboard_thread = threading.Thread(target=listen_keyboard, args=(main_window_hwnd, other_windows_hwnd))

    # Khởi chạy
    mouse_process.start()
    mouse_click_thread.start()
    keyboard_thread.start()

    # Chờ các luồng và process kết thúc
    mouse_click_thread.join()
    mouse_process.join()
    keyboard_thread.join()


if __name__ == "__main__":
    main()