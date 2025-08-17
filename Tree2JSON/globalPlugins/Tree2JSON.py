import appModuleHandler
import globalPluginHandler
from scriptHandler import script
import ui 
import api
import json
import win32api
import win32gui
from PIL import ImageGrab
import os
import time
from datetime import datetime

class GlobalPlugin(globalPluginHandler.GlobalPlugin):

    def __init__(self):
        super(GlobalPlugin, self).__init__()

    @script(
        description = "Capture screenshot and dump accessibility tree to JSON",
        category = "Accessibility Tree Dumper",
        gesture = "kb:NVDA+shift+p",
    )
    
    def script_dumpAccessibilityTree(self, gesture):
        """
        Main function to capture screenshot and dump accessibility data
        """
        try:
            # Create output directory
            output_dir = os.path.join(os.path.expanduser("~"), "Desktop", "nvda_dumps")
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)

            # Generate timestamp for unique filenames
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

            # Capture screenshot
            screenshot_path = self.capture_screenshot(output_dir, timestamp)

            # Get accessibility tree data
            accessibility_data = self.get_accessibility_tree()

            # Create JSON output
            json_output = self.create_json_output(screenshot_path, accessibility_data)

            # Save JSON file
            json_path = os.path.join(output_dir, f"accessibility_dump_{timestamp}.json")
            with open(json_path, 'w', encoding='utf-8') as f: json.dump(json_output, f, indent=2, ensure_ascii=False)

            ui.message(f"Screenshot and accessibility data saved to {output_dir}")

        except Exception as e:
            ui.message(f"Error: {str(e)}")

    def capture_screenshot(self, output_dir, timestamp):
        """
        Capture desktop screenshot
        """
        screenshot = ImageGrab.grab()
        screenshot_filename = f"screen_{timestamp}.png"
        screenshot_path = os.path.join(output_dir, screenshot_filename)
        screenshot.save(screenshot_path)
        return screenshot_filename # return relative path
    
    def get_accessibility_tree(self):
        """
        Extract accessibility tree data from NVDA
        """
        elements = []

        # get screen dimensions for normalization
        screen_width = win32api.GetSystemMetrics(0)
        screen_height = win32api.GetSystemMetrics(1)

        # Start from desktop object
        desktop = api.getDesktopObject()

        # recusively traverse the accessibility tree
        self.traverse_object(desktop, elements, screen_width, screen_height)

        return elements
    
    def traverse_object(self, obj, elements, screen_width, screen_height, max_depth=10, current_depth=0):
        """
        Recursively traverse accessibility object
        """
        if current_depth > max_depth:
            return
        
        # Extract element data from current object
        element_data = self.extract_element_data(obj, screen_width, screen_height)
        if element_data:
            elements.append(element_data)

        # Traverse children
        try:
            child = obj.firstChild
            while child:
                self.traverse_object(child, elements, screen_width, screen_height, max_depth, current_depth + 1)
                child = child.next
        except:
            pass

        def extract_element_data(self, obj, screen_width, screen_height):
            """
            Extract data from a single accessibility object
            """
            try:
                # Skip objects without location
                if not hasattr(obj, 'location') or not obj.location:
                    return None
                
                left, top, width, height = obj.location

                # skip objects with invalid dimensions
                if width <= 0 or height <= 0:
                    return None
                
                # Skip objects outside screen bounds
                if left < 0 or top < 0 or left >= screen_width or top >= screen_height:
                    return None
                
                # Normalized coordinates
                normalized_left = left / screen_width
                normalized_top = top / screen_height
                normalized_right = (left + width) / screen_width
                normalized_bottom = (top + height) / screen_height

                # calculate center point
                center_x = (normalized_left + normalized_right) / 2
                center_y  = (normalized_top + normalized_bottom)/ 2

                # Create instruction text
                instruction = self.create_instruction(obj)

                return {
                    "instruction": f'"{instruction}"',
                    "bbox": [normalized_left, normalized_top, normalized_right, normalized_bottom],
                    "point": [center_x, center_y]
                }
            
            except Exception as e:
                return None

        def create_instruction(self, obj):
            """
            Create descriptive instruction text for UI element
            """
            instruction_parts = []

            # Get name
            if hasattr(obj, 'name') and obj.name:
                instruction_parts.append(obj.name.strip())

            # Get role
            if hasattr(obj, 'role') and obj.role:
                try:
                    role_name = obj.role.displayString
                    if role_name and role_name not in instruction_parts:
                        instruction_parts.append(role_name.lower())
                except:
                    pass

            # Get description
            if hasattr(obj, 'description') and obj.description:
                desc = obj.description.strip()
                if desc and desc not in instruction_parts:
                    instruction_parts.append(desc)

            # get value for inputs
            if hasattr(obj, 'value') and obj.value:
                value = obj.value.strip()
                if value and len(value) < 50: # avoid very long values
                    instruction_parts.append(f"value: {value}")

            # Create final instruction
            if instruction_parts:
                instruction = " - ".join(instruction_parts)
            else:
                instruction = "UI element"

            # Limit length
            if len(instruction) > 100:
                instruction = instruction[:97] + "..."

            return instruction
        
        def create_json_output(self, screenshot_path, elements):
            """
            Create final JSON output in required format
            """
            # Get screenshot dimensions
            try:
                from PIL import Image
                with Image.open(os.path.join(os.path.dirname(screenshot_path), screenshot_path)) as img:
                    width, height = img.size
            except:
                # Fallback to screen dimensions
                width = win32api.GetSystemMetrics(0)
                height = win32api.GetSystemMetrics(1)

            return [{
                "img_url": screenshot_path,
                "img_size": [width, height],
                "element": elements,
                "element_size": len(elements)
            }]
