from crmc_med_phys.raystation.utils.misc_utils import get_current_or_abort


def show_nonempty_geoms():
    """Show (make visible) all ROIs and POIs with geometries on the current exam
    
    Hide (make invisible) all ROIs and POIs empty on the current exam
    """
    patient = get_current_or_abort('Patient')
    exam = get_current_or_abort('Examination')
    
    struct_set = case.PatientModel.StructureSets[exam.Name]

    # Show or hide each ROI
    with CompositeAction("Show/Hide ROIs"):
        for roi in case.PatientModel.RegionsOfInterest:
            visible = struct_set.RoiGeometries[roi.Name].HasContours()
            patient.SetRoiVisibility(RoiName=roi.Name, IsVisible=visible)

    # Show or hide each POI
    with CompositeAction("Show/Hide POIs"):
        for poi in case.PatientModel.PointsOfInterest:
            visible = abs(struct_set.PoiGeometries[poi.Name].Point.x) < 1000  # Infinite coordinates indicate empty geometry
            patient.SetPoiVisibility(PoiName=poi.Name, IsVisible=visible)


if __name__ == '__main__':
    show_nonempty_geoms()
