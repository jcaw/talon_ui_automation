from typing import Union, List, Optional, Iterator
import re
from queue import LifoQueue
from talon import Module, Context, actions, ui, cron

key = actions.key
sleep = actions.sleep


module = Module()
windows_context = Context()
windows_context.matches = "os: windows"
mac_context = Context()
mac_context.matches = "os: mac"


class Spec:
    def __init__(
        self,
        name: Optional[str] = None,
        class_name: Optional[str] = None,
        search_indirect: bool = False,
    ):
        self.name = name
        self.class_name = class_name
        self.search_indirect = search_indirect


def system_tray_button_spec(name_regexp: str) -> Spec:
    return Spec(
        name=name_regexp,
        class_name="SystemTray.NormalButton",
        search_indirect=True,
    )


class SearchSpecs:
    taskbar = [
        Spec(name="Taskbar", class_name="Shell_TrayWnd"),
    ]
    hidden_items_tray_button = [*taskbar, system_tray_button_spec("Show Hidden Icons")]
    overflow_tray = [Spec(class_name="TopLevelWindowForOverflowXamlIsland")]


def automator_find_elements(*search_specs: Spec) -> Iterator[ui.Element]:
    """Iterator to yeild all elements matching a particular search spec."""
    # TODO: Edge case for if the first spec matches the root node?

    windows = []
    browser_windows = []
    for window in ui.windows():
        if window.hidden or window.minimized:
            continue
        try:
            element = window.element
        except OSError:
            continue
        for browser in {"firefox", "chrome", "edge", "safari", "brave"}:
            if browser in element.name.lower():
                browser_windows.append(element)
                continue
        windows.append(element)
    # Browsers can take a long time to scrape, so put them at the end.
    windows.extend(browser_windows)

    queue = LifoQueue()
    for window in reversed(windows):
        queue.put((window, search_specs))

    while not queue.empty():
        element, remaining_specs = queue.get()
        if not remaining_specs:
            continue
        spec = remaining_specs[0]
        name_matches = spec.name is None or re.search(spec.name, element.name)
        class_matches = spec.class_name is None or re.search(
            spec.class_name, element.class_name
        )
        if name_matches and class_matches:
            if len(remaining_specs) == 1:
                yield element
            else:
                for child in element.children:
                    queue.put((child, remaining_specs[1:]))
        elif spec.search_indirect:
            # We want to search all intermediate nodes if search_indirect is set
            # - any unmatching node counts as a potential intermediate.
            for child in element.children:
                queue.put((child, remaining_specs))


class ElementNotFoundError(RuntimeError):
    pass


def automator_find_first_element(*search_specs: Spec) -> ui.Element:
    """Find the first element that matches `search_specs`."""
    try:
        return next(iter(automator_find_elements(*search_specs)))
    except StopIteration:
        raise ElementNotFoundError()


def click_element(element: ui.Element, button: int = 0):
    actions.mouse_move(*element.clickable_point)
    actions.mouse_click(button=button)


def automator_get_tray_icon(icon_name_regexp: str) -> ui.Element:
    # Open the start menu to ensure the tray is showing on Windows 11
    key("win")
    sleep("300ms")

    button_spec = system_tray_button_spec(icon_name_regexp)

    # Try to find the icon in the main window first.
    main_tray_button_specs = [*SearchSpecs.taskbar, button_spec]
    try:
        return automator_find_first_element(*main_tray_button_specs)
    except ElementNotFoundError:
        pass

    # If it's not in the main window, try finding it in the overflow icons.
    #
    # First open the overflow window.
    hidden_items_button = automator_find_first_element(
        *SearchSpecs.hidden_items_tray_button
    )
    click_element(hidden_items_button)
    sleep("100ms")

    # Once the overflow window is open,
    overflow_tray_button_spec = [*SearchSpecs.overflow_tray, button_spec]
    try:
        return automator_find_first_element(*overflow_tray_button_spec)
    except ElementNotFoundError:
        # Close the overflow tray (somewhat convoluted method to do so)
        key("win")
        sleep("200ms")
        click_element(hidden_items_button)
        sleep("200ms")
        key("win")
        raise ElementNotFoundError()


@module.action_class
class Actions:
    def automator_open_talon_repl():
        """Open the Talon repl from the menu (or switch to it if it's already open)."""

    def automator_open_talon_log():
        """Open the Talon log from the menu (or switch to it if it's already open)."""

    def automator_check_for_talon_updates():
        """Check for Talon updates."""

    def automator_click_tray_icon(icon_name_regexp: str, button: int = 0):
        """Click a tray icon on Windows."""


def exact_match_re(string: str) -> str:
    return f"^{re.escape(string)}$"


def automator_get_tray_icon_windows(icon_name_regexp: str) -> ui.Element:
    # Open the start menu to ensure the tray is showing on Windows 11
    key("win")
    sleep("300ms")

    button_spec = system_tray_button_spec(icon_name_regexp)

    # Try to find the icon in the main window first.
    main_tray_button_specs = [*SearchSpecs.taskbar, button_spec]
    try:
        return automator_find_first_element(*main_tray_button_specs)
    except ElementNotFoundError:
        pass

    # If it's not in the main window, try finding it in the overflow icons.
    #
    # First open the overflow window.
    hidden_items_button = automator_find_first_element(
        *SearchSpecs.hidden_items_tray_button
    )
    click_element(hidden_items_button)
    sleep("100ms")

    # Once the overflow window is open,
    overflow_tray_button_spec = [*SearchSpecs.overflow_tray, button_spec]
    try:
        return automator_find_first_element(*overflow_tray_button_spec)
    except ElementNotFoundError:
        # Close the overflow tray (somewhat convoluted method to do so)
        key("win")
        sleep("200ms")
        click_element(hidden_items_button)
        sleep("200ms")
        key("win")
        raise ElementNotFoundError()


def click_talon_menu_item_windows(*exact_menu_sequence: str):
    assert len(exact_menu_sequence) >= 1, exact_menu_sequence

    click_element(automator_get_tray_icon_windows(r"^Talon$"))
    sleep("100ms")

    # Require exact matches for menu items
    menu_path = [
        Spec(name="^Context$"),
        Spec(name=exact_match_re(exact_menu_sequence[0])),
    ]
    click_element(automator_find_first_element(*menu_path))
    if len(exact_menu_sequence) > 1:
        sleep("50ms")
        for i in range(1, len(exact_menu_sequence)):
            # The submenus appear to be named after the parent's name.
            subitem_path = [
                Spec(name=exact_match_re(exact_menu_sequence[i - 1])),
                Spec(name=exact_match_re(exact_menu_sequence[i])),
            ]
            click_element(automator_find_first_element(*subitem_path))


@windows_context.action_class("self")
class WindowsActions:
    def automator_open_talon_repl():
        # TODO: Switch to it if it's already open?
        click_talon_menu_item_windows("Scripting", "Console (REPL)")
        # Opening behaviour is a bit weird - unlike when the log is opened, it
        # doesn't close the start menu.
        sleep("2000ms")
        key("win")

    def automator_open_talon_log():
        # TODO: Switch to it if it's already open?
        click_talon_menu_item_windows("Scripting", "View Log")

    def automator_check_for_talon_updates():
        click_talon_menu_item_windows("Check for Updates...")

    def automator_click_tray_icon(icon_name_regexp: str, button: int = 0):
        click_element(automator_get_tray_icon_windows(icon_name_regexp), button=button)
