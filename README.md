# Med Phys Scripts
Code from my work as radiation Physicist Assistant at Cookeville Regional Medical Center

**Note:** See the [RayStation Scripts User Manual](./RayStation/RayStation Scripts User Manual.md) and [Other Scripts User Manual](./Other Scripts User Manual.md) for end-user instructions for running each script. This document is for technical details.
## How to Implement a Script in RayStation
See RaySearch's official Scripting Guideline for comprehensive coverage of how to implement a script in RayStation, but following is my clinic's system:

RayStation scripts are stored in `T:\Physics\KW\med-phys-scripts\RayStation`. Each "main" script defines a function with the same name as the file and calls this function inside `if __name__ == '__main__'`. The "blurb" scripts are in the `Blurbs` folder. Each "blurb" script has the same name as the "main" script prepended with an underscore, and simply runs the "main" script using RayStation's `connect.run` function. The "blurb" script is imported into RayStation. Separating the files in this way keeps us from having to invalidate, edit in RayStation or reimport from file, revalidate, and save the script every time we change the code.

We always use our custom RayStation scripting environment `CPython 3.8`.
### Example
Here is how we set up `add_box_to_external` in RayStation:
1. Write `add_box_to_external.py` and save in the `RayStation` folder.
    ```python
                ⋮
    def add_box_to_external():
                ⋮
    if __name__ == '__main__':
        add_box_to_external()
    ```
2. Write `_add_box_to_external.py` and save it in the `Blurbs` folder.
    ```python
    from connect import run

    run(r'T:\Physics\KW\Scripts\RayStation\add_box_to_external')
    ```
3. Import the blurb into RayStation and name the script `add_box_to_external`.