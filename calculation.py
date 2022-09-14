# -*- coding: utf-8 -*-
import clr # noqa
clr.AddReference("System")
from System.Collections.Generic import List # noqa

clr.AddReference('RevitServices')
from RevitServices.Persistence import DocumentManager # noqa
from RevitServices.Transactions import TransactionManager # noqa

clr.AddReference("RevitNodes")
import Revit # noqa
clr.ImportExtensions(Revit.Elements)
clr.ImportExtensions(Revit.GeometryConversion)

clr.AddReference('RevitAPI')
from Autodesk.Revit import DB # noqa
from Autodesk.Revit.DB import FilteredElementCollector as FEC # noqa
from Autodesk.Revit.DB import BuiltInCategory as BIC # noqa
from Autodesk.Revit.DB import ElementId # noqa

clr.AddReference("RevitAPIUI")
from Autodesk.Revit.UI import TaskDialog

import time


def get_rooms(view):
    msg = '\n\nСкрипт работает в двух режимах:'\
        '\n- запускается из спецификации (рассчитываются все '\
        'помещения спецификации);'\
        '\n- запускается из любого вида с предварительно выбранными '\
        'помещениями (рассчитываются только выбранные помещения).'
    if view.ViewType == DB.ViewType.Schedule:
        rooms = FEC(doc, view.Id).OfCategory(BIC.OST_Rooms)
        if rooms.GetElementCount()==0:
            msg = 'Расчёт не выполнен.\n\n'\
            'На текущей спецификации отсутствуют помещения для расчёта.' + msg
            TaskDialog.Show('Пустая спецификация', msg)
            rooms = []
    else:
        idd = [i.IntegerValue for i in uidoc.Selection.GetElementIds()]
        if idd:
            rooms = [doc.GetElement(ElementId(int(i))) for i in idd\
                    if doc.GetElement(ElementId(int(i))).Category.Name == "Помещения"]
        else:
            msg = 'Расчёт не выполнен.\n\n'\
                    'Ничего не выбрано.' + msg
            TaskDialog.Show('Ничего не выбрано', msg)
            rooms = []
    if rooms:
        rooms = [r for r in rooms if r.Area>0]
    return rooms

def delete_directshapes(rooms):
    global fin_report
    room_nums = []
    for room in rooms:
        room_nums.append(room.get_Parameter(DB.BuiltInParameter.ROOM_NUMBER).AsString())
    to_delete = []
    for el in FEC(doc).OfCategory(BIC.OST_Entourage).WhereElementIsNotElementType():
        ds_num = el.get_Parameter(DB.BuiltInParameter.ALL_MODEL_MARK).AsString()
        if ds_num in room_nums:
            to_delete.append(el.Id)
    fin_report += "Удалены лишние грани - " + str(len(to_delete)) + " шт. \n"
    TransactionManager.Instance.EnsureInTransaction(doc)
    doc.Delete(List[ElementId](to_delete))
    TransactionManager.Instance.ForceCloseTransaction()

def create_global_variables():
    global options
    global g_options
    global bb_tolerance
    global extrude_thin
    global height_plinth
    global directshapes_list
    global view_3d

    options = DB.SpatialElementBoundaryOptions()
    options.StoreFreeBoundaryFaces = True
    options.SpatialElementBoundaryLocation = DB.SpatialElementBoundaryLocation.Finish
    g_options = DB.Options()
    bb_tolerance = DB.UnitUtils.ConvertToInternalUnits(0.01, DB.UnitTypeId.Meters) #в метрах
    extrude_thin = DB.UnitUtils.ConvertToInternalUnits(0.001, DB.UnitTypeId.Meters) #в метрах
    height_plinth = DB.UnitUtils.ConvertToInternalUnits(0.05, DB.UnitTypeId.Meters)
    directshapes_list = []

    for v in FEC(doc).OfClass(DB.View):
        if v.ViewType == DB.ViewType.ThreeD and not v.IsTemplate:
            view_3d = v
            break
        else:
            view_3d = v

def get_wall_column_list_ids():
    all_walls = []
    for el in FEC(doc).OfClass(DB.Wall): # стены
        base = get_wall_column_base(el)
        if wall_exclude_keyword not in base.lower():
            if el.WallType.Kind != DB.WallKind.Curtain:
                all_walls.append(el.Id)
    for el in FEC(doc).OfCategory(BIC.OST_Walls).OfClass(DB.FamilyInstance): # стены - модель в контексте
        base = get_wall_column_base(el)
        if wall_exclude_keyword not in base.lower():
            all_walls.append(el.Id)
    for el in FEC(doc).OfCategory(BIC.OST_StructuralColumns).WhereElementIsNotElementType(): # несущие колонны
        base = get_wall_column_base(el)
        if wall_exclude_keyword not in base.lower():
            all_walls.append(el.Id)
    return List[DB.ElementId](all_walls)

def get_wall_column_base(el):
    if isinstance(el, DB.Wall):
        base = el.WallType.LookupParameter(wall_base_param_name).AsString()
    elif isinstance(el, DB.FamilyInstance):
        base = el.Symbol.LookupParameter(wall_base_param_name).AsString()
    return base if base else ""

def get_room_solids_by_sides_extrude(room, extrude_thin):
    solids_by_faces = []
    solid_fails = []
    for g in room.ClosedShell:
        for face in g.Faces:
            if isinstance(face, DB.PlanarFace):
                if face.FaceNormal.Z < 0.8 and face.FaceNormal.Z > -0.8:
                    curve_loop = face.GetEdgesAsCurveLoops()
                    try:
                        solids_by_faces.append(DB.GeometryCreationUtilities.CreateExtrusionGeometry(curve_loop, 
                            face.FaceNormal, 
                            extrude_thin))
                    except:
                        solid_fails.append(face.Area)
            elif isinstance(face, DB.CylindricalFace):
                try:
                    solids_by_faces.append(create_cylindric_solid(face))
                except:
                    solid_fails.append(face.Area)
    return solids_by_faces, solid_fails

def create_cylindric_solid(face, h=0):
    for edge_arr in face.EdgeLoops:
        curves = [i.AsCurve() for i in edge_arr]
        arcs = [i for i in curves if isinstance(i, DB.Arc)]
        lines = [i for i in curves if isinstance(i, DB.Line)]

        height = h if h else lines[0].Length
        norm_vect = DB.XYZ(0, 0, 1)

        arc = arcs[0] if arcs[0].GetEndPoint(0).Z < arcs[1].GetEndPoint(0).Z else arcs[1]
        arc_2 = arc.CreateOffset(extrude_thin, DB.XYZ(0, 0, -1))
        arc_2 = arc_2.CreateReversed()
        line_1 = DB.Line.CreateBound(arc.GetEndPoint(1), arc_2.GetEndPoint(0))
        line_2 = DB.Line.CreateBound(arc_2.GetEndPoint(1), arc.GetEndPoint(0))
        curve_loop = DB.CurveLoop()
        curve_loop.Append(arc)
        curve_loop.Append(line_1)
        curve_loop.Append(arc_2)
        curve_loop.Append(line_2)
        
        solid = DB.GeometryCreationUtilities.CreateExtrusionGeometry(List[DB.CurveLoop]([curve_loop]), norm_vect, height)
        return solid

def create_union_solid(solids):
    if solids:
        union = solids.pop()
        union_fails = 0
        if len(solids) > 0:
            for solid in solids:
                try:
                    union = DB.BooleanOperationsUtils.ExecuteBooleanOperation(union, solid, DB.BooleanOperationsType.Union)
                except:
                    union_fails += 1
    else:
        union, union_fails = [], 0
    return union, union_fails

def get_plinth_solids(room, height):
    up_vect = DB.XYZ(0, 0, height)
    solids = []
    solid_fails = []
    for g in room.ClosedShell:
        for face in g.Faces:
            if isinstance(face, DB.PlanarFace):
                if face.FaceNormal.Z < -0.99:
                    for curve_loop in face.GetEdgesAsCurveLoops():
                        for curve in curve_loop:
                            if hasattr(curve, "Direction"): # for line
                                vert_curve_loop = DB.CurveLoop()
                                p1 = curve.GetEndPoint(0)
                                p2 = curve.GetEndPoint(1)
                                p1_up = p1 + up_vect
                                p2_up = p2 + up_vect
                                vert_curve_loop.Append(curve)
                                vert_curve_loop.Append(DB.Line.CreateBound(p2, p2_up))
                                vert_curve_loop.Append(DB.Line.CreateBound(p2_up, p1_up))
                                vert_curve_loop.Append(DB.Line.CreateBound(p1_up, p1))
                                p_vect = (p2 - p1).Normalize()
                                norm_vect = DB.XYZ(p_vect.Y, p_vect.X, 0)
                                check_point = DB.Line.CreateBound(p2_up, p1_up).Evaluate(0.5, True) + norm_vect
                                if room.IsPointInRoom(check_point):
                                    norm_vect = norm_vect.Negate()
                                try:
                                    solids.append(DB.GeometryCreationUtilities.CreateExtrusionGeometry(
                                        List[DB.CurveLoop]([vert_curve_loop]),
                                        norm_vect,
                                        extrude_thin))
                                except:
                                    solid_fails.append(curve.Length)
            elif isinstance(face, DB.CylindricalFace):
                try:
                    solids.append(create_cylindric_solid(face, height))
                except:
                    pass
    return solids, solid_fails

def get_boundary_walls(room):
    room_bounds = []
    for segments in room.GetBoundarySegments(options):
        for segment in segments:
            element = doc.GetElement(segment.ElementId)
            if element:
                if element.Category.Id == ElementId(BIC.OST_Walls) or element.Category.Id == ElementId(BIC.OST_StructuralColumns):
                    base = get_wall_column_base(element)
                    if wall_exclude_keyword not in base.lower():
                        room_bounds.append(element.Id.IntegerValue)
    return list(set(room_bounds))

def get_around_walls(room):
    collector = FEC(doc, all_wall_ids)
    room_bb = room.BoundingBox[view_3d]
    outline = DB.Outline(room_bb.Min, room_bb.Max)
    bb_filter = DB.BoundingBoxIntersectsFilter(outline, bb_tolerance)
    wall_ids = collector.WherePasses(bb_filter).ToElementIds()
    return [i.IntegerValue for i in wall_ids]

def update_wall_solids(walls_id):
    for wall_id in walls_id:
        if wall_id not in walls_solids_flat:
            try:
                el_geometry = doc.GetElement(ElementId(wall_id)).Geometry[g_options]
                solids = []
                for geom in el_geometry:
                    if isinstance(geom, DB.Solid):
                        solids.append(geom)
                    elif isinstance(geom, DB.GeometryInstance):
                        for geom_instane in geom.GetInstanceGeometry():
                            if isinstance(geom_instane, DB.Solid):
                                solids.append(geom_instane)
                solid, fail_cnt = create_union_solid(solids)
                if fail_cnt:
                    test_out.append("several_solids_error-" + str(fail_cnt))
                walls_solids_flat[wall_id] = solid
            except:
                all_walls_error_id[wall_id] = "Error.Geometry"

def get_intersections(room_solids, walls_id):
    room_itersects = []
    walls_id_done = []
    wall_id_errors = []
    for wall_id in walls_id:
        for solid in room_solids:
            try:
                intersect = DB.BooleanOperationsUtils.ExecuteBooleanOperation(
                    solid,
                    walls_solids_flat[wall_id],
                    DB.BooleanOperationsType.Intersect)
                if intersect.Volume > DB.UnitUtils.ConvertToInternalUnits(1000, DB.UnitTypeId.CubicMillimeters):
                    room_itersects.append(intersect)
                    walls_id_done.append(wall_id)
            except:
                wall_id_errors.append(wall_id)
    return room_itersects, walls_id_done, wall_id_errors

def get_areas(room, intersections, walls_id_done):
    fin = room.LookupParameter(room_finish_param_name).AsString()
    fin = fin if fin else ""
    if len(intersections) == len(walls_id_done):
        mat_list = []
        area_list = []
        for wall_id, intersect in zip(walls_id_done, intersections):
            wall = doc.GetElement(ElementId(wall_id))
            base = get_wall_column_base(wall)
            mat_list.append(base + '.' + fin)
            i = intersect.Volume / extrude_thin
            area_list.append(DB.UnitUtils.ConvertFromInternalUnits(i, DB.UnitTypeId.SquareMeters))
        return area_list, mat_list

def get_area_strings(mat_list, area_list):
    py_faces_sum = 0
    mat_areas = {}
    for mat, area in zip(mat_list, area_list):
        if mat not in mat_areas.keys():
            mat_areas[mat] = area
        else:
            mat_areas[mat] += area
    area_list_2, mat_list_2 = [], []
    for mat in mat_areas.keys():
        mat_list_2.append(mat)
        area_list_2.append(mat_areas[mat])
        py_faces_sum += mat_areas[mat]
    wall_materials = ' : '.join(mat_list_2)
    wall_areas = ' : '.join([str(round(i, 2)) for i in area_list_2])
    return wall_materials, wall_areas, py_faces_sum

def update_directshape_list(room_num, intersections, area_list, mat_list, room_errors):
    for solid, area, mat in zip(intersections, area_list, mat_list):
        note = " : ".join([mat, str(area)])
        directshapes_list.append([solid, room_num, note, room_errors])

def create_directshape(solid, room_num, note, err_msg):
    ds = DB.DirectShape.CreateElement(doc, ElementId(BIC.OST_Entourage))
    geom = List[DB.GeometryObject]([solid])
    try:
        ds.SetShape(geom)
        ds.get_Parameter(DB.BuiltInParameter.ALL_MODEL_MARK).Set(room_num)
        ds.get_Parameter(DB.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(note)
        ds.LookupParameter("отделка.ошибки").Set(err_msg)
    except:
        ds_errors += 1

def get_room_errors(wall_materials, plinth_length, face_fails, plinth_fails, wall_id_errors):
    err = ""
    if wall_materials == "":
        err += "Стены не найдены или исключены!"
    elif plinth_length == 0:
        err += "Не удалось посчитать плинтус!"
    if face_fails:
        fails = DB.UnitUtils.ConvertFromInternalUnits(sum(face_fails), DB.UnitTypeId.SquareMeters)
        err += 'Ошибки_грани - ' + str(round(fails, 2)) + ' м2; '
    if plinth_fails:
        fails = DB.UnitUtils.ConvertFromInternalUnits(sum(plinth_fails), DB.UnitTypeId.Meters)
        err += 'Ошибки_плинтус - ' + str(round(fails, 2)) + 'м. п.; '
    if wall_id_errors:
        temp = []
        for wall_id in wall_id_errors:
            temp.append(get_wall_column_base(doc.GetElement(DB.ElementId(wall_id))))
        err += "Ошибки_стены - " + "; ".join(list(set(temp))) + "; "
    return err

def update_wall_errors(room_num, wall_id_errors):
    for wall_id in wall_id_errors:
        if wall_id not in all_walls_error_id.keys():
            all_walls_error_id[wall_id] = "Error.Rooms - " + room_num
        else:
            all_walls_error_id[wall_id] += "; " + room_num

def get_floor_in_doors(room):
    pass

def get_stairs_info():
    pass

doc = DocumentManager.Instance.CurrentDBDocument
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
active_view = doc.ActiveView
fin_report = ""
script_start_time = time.time()

room_finish_param_name = "отделка.стены.финиш"
wall_base_param_name = "отделка.стены.основа"
wall_exclude_keyword = "исключить"

rooms = get_rooms(active_view)
fin_report += "Помещений посчитано - " + str(len(rooms)) + "\n\n"
if rooms:
    delete_directshapes(rooms)
    create_global_variables()
    all_wall_ids = get_wall_column_list_ids()

    all_walls_error_id = dict()
    walls_solids_flat = dict()
    test_out = []
    to_set_parameters = []
    ds_errors = 0
    core_loop_start_time = time.time()

    for room in rooms: # core loop
        start_time = time.time()
        room_errors, room_info = "", ""
        room_num = room.get_Parameter(DB.BuiltInParameter.ROOM_NUMBER).AsString()

        room_solids, face_fails = get_room_solids_by_sides_extrude(room, extrude_thin) # стены
        walls_id = get_boundary_walls(room)
        walls_id_around = get_around_walls(room)
        walls_id_around = [i for i in walls_id_around if i not in walls_id]
        update_wall_solids(walls_id + walls_id_around)
        intersections, walls_id_done, wall_id_errors = get_intersections(room_solids, walls_id)
        intersections_around, walls_id_done_around, temp = get_intersections(room_solids, walls_id_around)
        intersections += intersections_around
        walls_id_done += walls_id_done_around
        area_list, mat_list = get_areas(room, intersections, walls_id_done)
        wall_materials, wall_areas, py_faces_sum = get_area_strings(mat_list, area_list)

        room_solid_plinth, plinth_fails = get_plinth_solids(room, height_plinth) # плинтусы
        plinth_intersections, temp, temp = get_intersections(room_solid_plinth, walls_id + walls_id_around)
        plinth_length = 0
        for plinth_s in plinth_intersections:
            plinth_length += plinth_s.Volume / extrude_thin / height_plinth
        plinth_length = DB.UnitUtils.ConvertFromInternalUnits(round(plinth_length, 2), DB.UnitTypeId.Meters)

        room_errors += get_room_errors(wall_materials, plinth_length, face_fails, plinth_fails, wall_id_errors)
        update_wall_errors(room_num, wall_id_errors)

        update_directshape_list(room_num, intersections, area_list, mat_list, room_errors)
        update_directshape_list(
            room_num, \
            plinth_intersections, \
            [str(plinth_length)] * len(plinth_intersections), \
            ['плинтус'] * len(plinth_intersections), \
            room_errors)

        floor_in_doors = get_floor_in_doors(room)

        stairs_info = get_stairs_info()

        duration = time.time() - start_time
        room_info += "расчет_секунд-" + str(duration)
        to_set_parameters.append([room, wall_materials, wall_areas, plinth_length, room_errors, room_info, py_faces_sum])

    core_loop_duration = time.time() - core_loop_start_time
    test_out.append(core_loop_duration)

    TransactionManager.Instance.EnsureInTransaction(doc)

    for room, wall_materials, wall_areas, plinth_length, room_errors, room_info, py_faces_sum in to_set_parameters:
        room.LookupParameter("отделка.грани.материал").Set(wall_materials)
        room.LookupParameter("отделка.грани.площадь").Set(wall_areas)
        room.LookupParameter("отделка.плинтус.длина").Set(plinth_length)
        room.LookupParameter("отделка.ошибки").Set(room_errors)
        room.LookupParameter("py_faces_sum").Set(round(py_faces_sum, 2))

    for wall_id in all_walls_error_id.keys():
        wall = doc.GetElement(DB.ElementId(wall_id))
        wall.LookupParameter("отделка.ошибки").Set(all_walls_error_id[wall_id])

    for solid, room_num, note, room_errors in directshapes_list:
        create_directshape(solid, room_num, note, room_errors)

    TransactionManager.Instance.ForceCloseTransaction()

    fin_report += "Создано новых граней - " + str(len(directshapes_list)) + " шт. \n"
    fin_report += "Время выполнения скрипта - " + str(round(time.time() - script_start_time, 2)) + " сек. \n"
    TaskDialog.Show('Отчет', fin_report)

