# NVDAPlugins
A repo of useful NVDA plugins

## Tree2JSON

### To save the plugin 

1. Save the `Tree2JSON.py` file in your NVDA user configuration directory
2. Restart NVDA: The plugin will be loaded automatically
3. Use the hotkey: Press `NVDA+Shift+P` to capture a screenshot and dump accessibility data

### Output format
The plugin generates the follow JSON (the following is an example output)
```
[
  {
    "img_url": "screen_20250817_105030.png",
    "img_size": [1920, 1080],
    "element": [
      {
        "instruction": "\"Start button - button\"",
        "bbox": [0.0, 0.9537037037037037, 0.0416666666666667, 1.0],
        "point": [0.0208333333333333, 0.9768518518518519]
      }
    ],
    "element_size": 1
  }
]

```
