# RayStation Scripts User Manual
This is a how-to guide for end users of Kaley's RayStation Python scripts. If the script you need is not yet imported into your RayStation, see this repository's [README](./README.md). For scripts that are not for RayStation, see the [Other Scripts User Manual](./Other Scripts User Manual.md).
## add_box_to_external
Modifies the external (body) geometry on the currently loaded exam, to include some of the Vac-Lok&trade; geometry used for SBRT patients.

SBRT lung patients are positioned using a Vac-Lok&trade; bag. When planning, we prefer our external geometry to include the body of the patient as well as some area posterior and right/left of the patient. This area ideally encompasses the entire couch geometry in the I-S and R-L directions, and up to the localization point in the P-A direction. In practice, we crop the area by a pixel on each R-L side to prevent possible calculation errors due to the external extending outside the image.

Of course, we want to retain the original external geometry, so before adding the box to it, we copy the external geometry into a new ROI called External^NoBox.
### Example
here is what the external looks like before running the script; it's just the body:<br>
<img src="./images_for_readme/add_box_to_external/external_before_script.png" alt="External geometry before running the script" height="300"/><br>
After running the script, we have the modified external:<br>
<img src="./images_for_readme/add_box_to_external/external_after_script.png" alt="External geometry after running the script" height="300"/><br>
and External^NoBox, containing the original external geometry:<br>
<img src="./images_for_readme/add_box_to_external/external^no_box_after_script.png" alt="External^NoBox" height="300"/>
<hr>

## copy_plan_without_changes
RayStation's **Copy plan** functionality is amazing, but we wanted to change a couple of naming conventions in the copy:
<ul>
    <li>Retain beam set and isocenter names.
        <p>The beam sets and beam isocenters in the copied plan have the same name as the plan, plus a copy number. For example, if a plan called <code>R Breast</code> with two beam sets is copied to a plan called <code>R Breast (1)</code>, the new beam sets are called <code>R Breast (1)</code> and <code>R Breast (1)_2</code>respectively, regardless of their names in <code>R Breast</code>. Likewise, a beam isocenters are called <code>R Breast (1) 1</code>, <code>R Breast (1) 2</code>, etc.</p>
    </li>
    <li>Renumber the new beams.
        <p>We make the new beam numbers, including setup beam numbers, unique among all beam numbers in the patient. The first beam is numbered one more than the greatest beam number in the patient, and beam numbers increase consecutively from there.</p>
    </li>
    <li>Name beams the same as their numbers.
        <p>We want to retain any important information from the old name, though, so if an old beam name is different from its number (or a setup beam name is not the default), we append the old name to the new description.</p> 
    </li>
</ul>

### Example
Given a plan `Plan 1`:
<p float="left">
    <img src="./images_for_readme/copy_plan_without_changes/old_plan_beam_sets.png" alt="Plan 1 beam sets" height="200"/>
    <img src="./images_for_readme/copy_plan_without_changes/old_plan_beams.png" alt="Plan 1 beams" height="200"/>
</p>
If we copy `Plan 1` using the **Copy plan** button, naming the copy `Plan 1 (1)`, we get these names and numbers:
<p float="left">
    <img src="./images_for_readme/copy_plan_without_changes/new_plan_beam_sets_no_script.png" alt="Plan 1 (1) beam sets, without script" height="200"/>
    <img src="./images_for_readme/copy_plan_without_changes/new_plan_beams_no_script.png" alt="Plan 1 (1) beams, without script" height="200"/>
</p>
If we instead copy the plan using the script, the `Plan 1 (1)` beam sets and beams match those in `Plan 1` and are uniquely numbered:
<p float="left">
    <img src="./images_for_readme/copy_plan_without_changes/new_plan_beam_sets_with_script.png" alt="Plan 1 (1) beam sets, without script" height="200"/>
    <img src="./images_for_readme/copy_plan_without_changes/new_plan_beams_with_script.png" alt="Plan 1 (1) beams, without script" height="200"/>
</p>