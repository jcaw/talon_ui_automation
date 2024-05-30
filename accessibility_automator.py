from typing import Union, List, Optional, Iterator
import re
from queue import LifoQueue
from talon import Module, Context, actions, ui, cron, app, canvas
from talon.ui import Rect, Point2d

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


@module.action
def automator_spec(
    name: Optional[str] = None,
    class_name: Optional[str] = None,
    search_indirect: Optional[bool] = False,
) -> Spec:
    """Create and return an automation element `Spec` object."""
    return Spec(name=name, class_name=class_name, search_indirect=search_indirect)


def system_tray_button_spec(name_regexp: str) -> Spec:
    return Spec(
        name=name_regexp,
        class_name="SystemTray.NormalButton",
        search_indirect=True,
    )


class SearchSpecs:
    """Useful predefined search paths."""

    WINDOWS_TASKBAR = [
        Spec(name="Taskbar", class_name="Shell_TrayWnd"),
    ]
    WINDOWS_HIDDEN_ITEMS_TRAY_BUTTON = [
        *WINDOWS_TASKBAR,
        system_tray_button_spec("Show Hidden Icons"),
    ]
    WINDOWS_TRAY_ICONS_OVERFLOW = [
        Spec(class_name="TopLevelWindowForOverflowXamlIsland")
    ]


@module.action
def automator_predefined_specs() -> SearchSpecs:
    """Get the predefined search specs object."""
    return SearchSpecs


def draw(c: canvas.Canvas):
    TRANSPARENCY = "77"

    paint = c.paint
    paint.blendmode = paint.Blend.SRC
    paint.color = "#000000" + TRANSPARENCY
    paint.style = paint.Style.FILL
    c.draw_rect(c.rect)

    paint.textsize = round(min(c.rect.width, c.rect.height) / 8)
    paint.color = "#FFFFFF" + TRANSPARENCY
    text = "(Automating UI)"
    text_dims = paint.measure_text(text)[1]
    c.draw_text(
        text,
        c.rect.center.x - text_dims.width / 2,
        # HACK: Compensate for the fact the text's y is measured from
        #  the line base, not the tails of the text.
        c.rect.center.y + text_dims.height / 3,
    )


canvases = []


def create_canvases():
    if not canvases:
        print("Creating canvases")
        for screen in ui.screens():
            c = canvas.Canvas.from_screen(screen)
            # HOTFIX: from_screen not working right on Windows
            if app.platform == "windows":
                hotfix_rect = Rect(*screen.rect)
                hotfix_rect.height -= 1
                c.rect = hotfix_rect
            c.focusable = False
            c.register("draw", draw)
            c.freeze()
            canvases.append(c)


def destroy_canvases():
    for c in canvases:
        c.unregister("draw", draw)
        c.close()
    canvases.clear()


canvas_context_count = 0


# TODO: Convert this to an action to remove need to import?
class AutomationOverlay:
    def __enter__(self):
        global canvas_context_count
        # Count multiple entries into this context so the canvases are only
        # destroyed when exiting the outermost context.
        if canvas_context_count == 0:
            create_canvases()
        canvas_context_count += 1
        return self

    def __exit__(self, *_, **__):
        global canvas_context_count
        canvas_context_count -= 1
        if canvas_context_count == 0:
            destroy_canvases()
        return False


@module.action
def automator_overlay() -> AutomationOverlay:
    """Get a context manager that creates an automation overlay."""
    return AutomationOverlay()


def automator_find_elements_from_roots(
    root_elements: List[ui.Element], *search_specs: Spec
):
    queue = LifoQueue()
    for element in reversed(root_elements):
        queue.put((element, search_specs))

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
    return automator_find_elements_from_roots(reversed(windows), *search_specs)


def automator_find_elements_current_window(*search_specs: Spec) -> Iterator[ui.Element]:
    return automator_find_elements_from_roots(
        [ui.current_window().element], *search_specs
    )


class ElementNotFoundError(RuntimeError):
    pass


def _automator_find_first_element_internal(
    elements_iterator, search_specs
) -> ui.Element:
    """Common functionality. See references."""
    try:
        return next(iter(elements_iterator))
    except StopIteration:
        raise ElementNotFoundError()


def automator_find_first_element(*search_specs: Spec) -> ui.Element:
    """Find the first element that matches `search_specs`."""
    return _automator_find_first_element_internal(
        automator_find_elements(*search_specs)
    )


def automator_find_first_element_current_window(*search_specs: Spec) -> ui.Element:
    """Find the first element that matches `search_specs` in the current window."""
    return _automator_find_first_element_internal(
        automator_find_elements_current_window(*search_specs)
    )


def click_element(element: ui.Element, button: int = 0):
    # TODO: Return mouse to original position?
    actions.mouse_move(*element.clickable_point)
    actions.mouse_click(button=button)


def automator_get_tray_icon(icon_name_regexp: str) -> ui.Element:
    # Open the start menu to ensure the tray is showing on Windows 11
    key("win")
    sleep("300ms")

    button_spec = system_tray_button_spec(icon_name_regexp)

    # Try to find the icon in the main window first.
    main_tray_button_specs = [*SearchSpecs.WINDOWS_TASKBAR, button_spec]
    try:
        return automator_find_first_element(*main_tray_button_specs)
    except ElementNotFoundError:
        pass

    # If it's not in the main window, try finding it in the overflow icons.
    #
    # First open the overflow window.
    hidden_items_button = automator_find_first_element(
        *SearchSpecs.WINDOWS_HIDDEN_ITEMS_TRAY_BUTTON
    )
    click_element(hidden_items_button)
    sleep("100ms")

    # Once the overflow window is open,
    overflow_tray_button_spec = [*SearchSpecs.WINDOWS_TRAY_ICONS_OVERFLOW, button_spec]
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
    def automator_click_element(search_specs: List[Spec], button: int = 0):
        """Find and click a specific element."""
        with AutomationOverlay():
            click_element(automator_find_first_element(*search_specs), button=button)

    def automator_click_element_current_window(
        search_specs: List[Spec], button: int = 0
    ):
        """Find and click a specific element in the current window."""
        with AutomationOverlay():
            # TODO: UI automation, click element in the current window.
            click_element(
                automator_find_first_element_current_window(search_specs), button=button
            )

    def automator_open_talon_repl():
        """Open the Talon repl from the menu (or switch to it if it's already open)."""

    def automator_open_talon_log():
        """Open the Talon log from the menu (or switch to it if it's already open)."""

    def automator_check_for_talon_updates():
        """Check for Talon updates."""

    # TODO: Use same action on Mac? Does it have the same concept of a tray?
    def automator_click_tray_icon(icon_name_regexp: str, button: int = 0):
        """Click a tray icon on Windows."""

    def automator_close_start_menu():
        """Close the start menu in Windows, iff it's open. Does nothing on other platforms."""
        if app.platform == "windows":
            with AutomationOverlay():
                for window in ui.windows():
                    if (
                        not (window.hidden or window.minimized)
                        and window.title == "Start"
                        and window.app.name == "Windows Start Experience Host"
                    ):
                        print(
                            "[accessibility_automator]: Start menu detected as open. Closing it."
                        )
                        key("win")
                        sleep("500ms")
                        return


def exact_match_re(string: str) -> str:
    return f"^{re.escape(string)}$"


def automator_get_tray_icon_windows(icon_name_regexp: str) -> ui.Element:
    # Reset so we have a predictable starting state
    actions.self.automator_close_start_menu()

    # Open the start menu to ensure the tray is showing on Windows 11
    key("win")
    sleep("300ms")

    button_spec = system_tray_button_spec(icon_name_regexp)

    # Try to find the icon in the main window first.
    main_tray_button_specs = [*SearchSpecs.WINDOWS_TASKBAR, button_spec]
    try:
        return automator_find_first_element(*main_tray_button_specs)
    except ElementNotFoundError:
        pass

    # If it's not in the main window, try finding it in the overflow icons.
    #
    # First open the overflow window.
    hidden_items_button = automator_find_first_element(
        *SearchSpecs.WINDOWS_HIDDEN_ITEMS_TRAY_BUTTON
    )
    click_element(hidden_items_button)
    sleep("100ms")

    # Once the overflow window is open,
    overflow_tray_button_spec = [*SearchSpecs.WINDOWS_TRAY_ICONS_OVERFLOW, button_spec]
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
        with AutomationOverlay():
            # TODO: Switch to it if it's already open?
            click_talon_menu_item_windows("Scripting", "Console (REPL)")
            # Opening behaviour is a bit weird - unlike when the log is opened, it
            # doesn't close the start menu.
            sleep("2000ms")
            key("win")

    def automator_open_talon_log():
        with AutomationOverlay():
            # TODO: Switch to it if it's already open?
            click_talon_menu_item_windows("Scripting", "View Log")

    def automator_check_for_talon_updates():
        with AutomationOverlay():
            click_talon_menu_item_windows("Check for Updates...")

    def automator_click_tray_icon(icon_name_regexp: str, button: int = 0):
        with AutomationOverlay():
            click_element(
                automator_get_tray_icon_windows(icon_name_regexp), button=button
            )
