"""Tests for the COM backend using fake AutoCAD COM objects.

No live AutoCAD and no win32com required: the backend imports pywin32
lazily, and these tests monkeypatch ``_load_com`` plus the VARIANT
helpers.
"""

from __future__ import annotations

import asyncio
import math
import threading
import time
import types

import pytest

import autocad_mcp.backends.com_backend as com_backend


class FakeComError(Exception):
    def __init__(self, hresult, message="fake COM error"):
        super().__init__(hresult, message)


# ---------------------------------------------------------------------------
# Fake COM object model
# ---------------------------------------------------------------------------


class FakeEntity:
    _counter = 0

    def __init__(self, entity_name="AcDbLine", layer="0", **props):
        FakeEntity._counter += 1
        self.Handle = format(FakeEntity._counter, "X")
        self.Layer = layer
        self.EntityName = entity_name
        self.deleted = False
        self.moved = []
        self.rotated = None
        self.scaled = None
        self.attrs = []
        self.__dict__.update(props)

    def Delete(self):
        self.deleted = True

    def Move(self, p1, p2):
        self.moved.append((tuple(p1), tuple(p2)))

    def Rotate(self, center, angle):
        self.rotated = (tuple(center), angle)

    def ScaleEntity(self, center, factor):
        self.scaled = (tuple(center), factor)

    def Mirror(self, p1, p2):
        return FakeEntity(self.EntityName, self.Layer)

    def Copy(self):
        return FakeEntity(self.EntityName, self.Layer)

    def Offset(self, distance):
        return (FakeEntity(self.EntityName, self.Layer),)

    def ArrayRectangular(self, *args):
        self.array_args = args

    def GetAttributes(self):
        return self.attrs


class FakeModelSpace:
    def __init__(self):
        self.entities = []
        self.calls = []
        self.thread_ids = set()
        self.fail_active_calls = 0
        self.fail_hresult = 0x80010001
        self.delay = 0.0

    @property
    def Count(self):
        return len(self.entities)

    def __iter__(self):
        return iter(self.entities)

    def _record(self, name, *args):
        self.thread_ids.add(threading.get_ident())
        self.calls.append((name, args))
        if self.fail_active_calls > 0:
            self.fail_active_calls -= 1
            raise FakeComError(self.fail_hresult)
        if self.delay:
            time.sleep(self.delay)

    def _add(self, entity_name, **props):
        entity = FakeEntity(entity_name, **props)
        self.entities.append(entity)
        return entity

    def AddLine(self, p1, p2):
        self._record("AddLine", tuple(p1), tuple(p2))
        return self._add("AcDbLine", StartPoint=p1, EndPoint=p2)

    def AddCircle(self, center, radius):
        self._record("AddCircle", tuple(center), radius)
        return self._add("AcDbCircle", Center=center, Radius=radius)

    def AddLightWeightPolyline(self, points):
        self._record("AddLightWeightPolyline", tuple(points))
        entity = self._add("AcDbPolyline")
        entity.points = list(points)
        entity.Closed = False
        return entity

    def AddArc(self, center, radius, start, end):
        self._record("AddArc", tuple(center), radius, start, end)
        return self._add("AcDbArc")

    def AddEllipse(self, center, major_axis, ratio):
        self._record("AddEllipse", tuple(center), tuple(major_axis), ratio)
        return self._add("AcDbEllipse")

    def AddText(self, text, point, height):
        self._record("AddText", text, tuple(point), height)
        return self._add("AcDbText", TextString=text, Rotation=0.0)

    def AddMText(self, point, width, text):
        self._record("AddMText", tuple(point), width, text)
        return self._add("AcDbMText", Height=2.5)

    def AddHatch(self, pattern_type, pattern_name, associative):
        self._record("AddHatch", pattern_type, pattern_name, associative)
        entity = self._add("AcDbHatch")

        def append_outer_loop(objects):
            entity.boundary = objects

        entity.AppendOuterLoop = append_outer_loop
        entity.Evaluate = lambda: None
        return entity

    def AddDimLinear(self, p1, p2, location):
        self._record("AddDimLinear", tuple(p1), tuple(p2), tuple(location))
        return self._add("AcDbRotatedDimension")

    def AddDimRotated(self, p1, p2, location, rotation):
        self._record("AddDimRotated", tuple(p1), tuple(p2), tuple(location), rotation)
        return self._add("AcDbRotatedDimension")

    def AddDimAligned(self, p1, p2, location):
        self._record("AddDimAligned", tuple(p1), tuple(p2), tuple(location))
        return self._add("AcDbAlignedDimension")

    def AddDimAngular(self, center, p1, p2, location):
        self._record("AddDimAngular", tuple(center), tuple(p1), tuple(p2), tuple(location))
        return self._add("AcDbAngularDimension2")

    def AddDimRadial(self, center, chord, leader_length):
        self._record("AddDimRadial", tuple(center), tuple(chord), leader_length)
        return self._add("AcDbRadialDimension")

    def AddDimDiametric(self, p1, p2, leader_length):
        self._record("AddDimDiametric", tuple(p1), tuple(p2), leader_length)
        return self._add("AcDbDiametricDimension")

    def AddLeader(self, points):
        self._record("AddLeader", tuple(points))
        return self._add("AcDbLeader")

    def AppendEntity(self, entity):
        self.entities.append(entity)

    def InsertBlock(self, point, name, xs, ys, zs, rotation):
        self._record("InsertBlock", tuple(point), name, xs, ys, zs, rotation)
        return self._add("AcDbBlockReference", Name=name)


class FakeLayer:
    def __init__(self, name, color=7):
        self.Name = name
        self.Color = color
        self.Linetype = "Continuous"
        self.LayerOn = True
        self.Freeze = False
        self.Lock = False
        self.Lineweight = 25


class FakeLayers:
    def __init__(self):
        self.items = {}

    def Item(self, name):
        try:
            return self.items[name]
        except KeyError:
            raise FakeComError(0x80004005, f"layer {name} not found")

    def Add(self, name):
        if name in self.items:
            raise FakeComError(0x80004005, "duplicate")
        layer = FakeLayer(name)
        self.items[name] = layer
        return layer

    def __iter__(self):
        return iter(self.items.values())


class FakeBlock:
    def __init__(self, name):
        self.Name = name
        self.entities = []

    def AddLine(self, p1, p2):
        entity = FakeEntity("AcDbLine")
        self.entities.append(entity)
        return entity

    def AddCircle(self, center, radius):
        entity = FakeEntity("AcDbCircle")
        self.entities.append(entity)
        return entity

    def AddText(self, text, point, height):
        entity = FakeEntity("AcDbText")
        self.entities.append(entity)
        return entity


class FakeBlocks:
    def __init__(self):
        self.items = {"*Model_Space": FakeBlock("*Model_Space"), "*Paper_Space": FakeBlock("*Paper_Space")}

    def Add(self, point, name):
        block = FakeBlock(name)
        self.items[name] = block
        return block

    def __iter__(self):
        return iter(self.items.values())


class FakePlot:
    def __init__(self):
        self.QuietErrorMode = False
        self.plot_file = None

    def PlotToFile(self, path):
        self.plot_file = path


class FakeDocument:
    def __init__(self, name="Drawing1.dwg", full_name=""):
        self.Name = name
        self.FullName = full_name
        self.ModelSpace = FakeModelSpace()
        self.Layers = FakeLayers()
        self.Blocks = FakeBlocks()
        self.ActiveLayer = None
        self.Plot = FakePlot()
        self._variables = {"FILEDIA": 1}
        self.commands = []
        self.saved = []
        self.saved_as = []

    def HandleToObject(self, handle):
        for entity in self.ModelSpace.entities:
            if entity.Handle.upper() == str(handle).upper():
                return entity
        raise FakeComError(0x80004005, "handle not found")

    def GetVariable(self, name):
        return self._variables.get(name)

    def SetVariable(self, name, value):
        self._variables[name] = value

    def Save(self):
        self.saved.append(self.FullName)

    def SaveAs(self, path):
        self.saved_as.append(path)
        self.FullName = path

    def PurgeAll(self):
        self.purged = True

    def SendCommand(self, command):
        self.commands.append(command)


class FakeDocuments:
    def __init__(self, app):
        self.app = app
        self.docs = []

    @property
    def Count(self):
        return len(self.docs)

    def Add(self):
        doc = FakeDocument(f"Drawing{len(self.docs) + 1}.dwg")
        self.docs.append(doc)
        self.app._active = doc
        return doc

    def Open(self, path):
        doc = FakeDocument(path.rsplit("\\", 1)[-1], path)
        self.docs.append(doc)
        self.app._active = doc
        return doc

    def Item(self, index):
        return self.docs[index]


class FakeApplication:
    def __init__(self, documents=1):
        self.Version = "26.0s (Fake)"
        self.HWND = 4242
        self.Visible = False
        self.Documents = FakeDocuments(self)
        self._active = None
        self.zoom_calls = []
        self.fail_active_once = None
        self.disconnected = False
        for _ in range(documents):
            self.Documents.Add()

    @property
    def ActiveDocument(self):
        if self.fail_active_once is not None:
            error = FakeComError(self.fail_active_once)
            self.fail_active_once = None
            raise error
        if self.disconnected:
            raise FakeComError(0x80010108, "RPC_E_DISCONNECTED")
        if self._active is None:
            raise FakeComError(0x80004005, "no active document")
        return self._active

    def ZoomExtents(self):
        self.zoom_calls.append("extents")

    def ZoomWindow(self, p1, p2):
        self.zoom_calls.append(("window", tuple(p1), tuple(p2)))


class FakeClient:
    """Stands in for win32com.client."""

    def __init__(self, app, get_active_error=None):
        self.app = app
        self.get_active_error = get_active_error
        self.calls = []

    def GetActiveObject(self, progid):
        self.calls.append(("GetActiveObject", progid))
        if self.get_active_error is not None:
            raise self.get_active_error
        self.app.disconnected = False
        return self.app

    def Dispatch(self, progid):
        self.calls.append(("Dispatch", progid))
        return self.app


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------


@pytest.fixture
async def backend_factory(monkeypatch):
    """Build ComBackend instances wired to fake COM objects."""
    created = []

    def make(app, get_active_error=None, autolaunch=True, call_timeout=5.0):
        client = FakeClient(app, get_active_error)
        fake = {
            "pythoncom": types.SimpleNamespace(
                CoInitialize=lambda: None, CoUninitialize=lambda: None
            ),
            "client": client,
            "pywintypes": types.SimpleNamespace(com_error=FakeComError),
        }
        monkeypatch.setattr(com_backend, "_load_com", lambda: fake)
        monkeypatch.setattr(com_backend, "_r8", lambda values: list(values))
        monkeypatch.setattr(com_backend, "_dispatch_array", lambda objects: tuple(objects))

        backend = com_backend.ComBackend(autolaunch=autolaunch)
        backend._call_timeout = call_timeout
        backend._test_client = client
        created.append(backend)
        return backend

    yield make

    for backend in created:
        try:
            await backend.close()
        except Exception:
            pass


# ---------------------------------------------------------------------------
# Connection lifecycle
# ---------------------------------------------------------------------------


async def test_initialize_attaches_to_running_instance(backend_factory):
    app = FakeApplication(documents=1)
    backend = backend_factory(app)

    result = await backend.initialize()
    assert result.ok
    assert result.payload["backend"] == "com"
    assert result.payload["document"] == "Drawing1.dwg"

    status = await backend.status()
    assert status.ok
    assert status.payload["backend"] == "com"
    assert status.payload["connected"] is True
    assert status.payload["active_document"]["name"] == "Drawing1.dwg"


async def test_initialize_launches_autocad_when_not_running(backend_factory):
    app = FakeApplication(documents=1)
    backend = backend_factory(app, get_active_error=FakeComError(0x800401E3))

    result = await backend.initialize()
    assert result.ok
    # Dispatch was used because GetActiveObject failed.
    assert ("Dispatch", "AutoCAD.Application") in backend._test_client.calls


async def test_initialize_does_not_launch_when_disabled(backend_factory):
    app = FakeApplication(documents=1)
    backend = backend_factory(app, get_active_error=FakeComError(0x800401E3), autolaunch=False)

    result = await backend.initialize()
    assert result.ok is False
    assert "not running" in result.error


async def test_creates_document_when_none_open(backend_factory):
    app = FakeApplication(documents=0)
    backend = backend_factory(app)

    result = await backend.initialize()
    assert result.ok
    assert app.Documents.Count == 1


async def test_calls_run_on_a_single_background_thread(backend_factory):
    app = FakeApplication()
    backend = backend_factory(app)
    assert (await backend.initialize()).ok

    await backend.create_line(0, 0, 1, 1)
    await backend.create_circle(0, 0, 5)

    thread_ids = app._active.ModelSpace.thread_ids
    assert len(thread_ids) == 1
    assert threading.get_ident() not in thread_ids


async def test_reconnects_once_when_autocad_disconnects(backend_factory):
    app = FakeApplication()
    backend = backend_factory(app)
    assert (await backend.initialize()).ok

    app.disconnected = True
    result = await backend.create_line(0, 0, 10, 10)

    assert result.ok
    assert backend._reconnect_count == 1


# ---------------------------------------------------------------------------
# Drawing / entities
# ---------------------------------------------------------------------------


async def test_rectangle_is_closed_four_point_polyline(backend_factory):
    app = FakeApplication()
    backend = backend_factory(app)
    await backend.initialize()

    result = await backend.create_rectangle(0, 0, 100, 60, layer="PART_DEMO")
    assert result.ok
    assert result.payload["entity_type"] == "LWPOLYLINE"

    entity = app._active.ModelSpace.entities[-1]
    assert entity.Closed is True
    assert entity.points == [0.0, 0.0, 100.0, 0.0, 100.0, 60.0, 0.0, 60.0]
    assert entity.Layer == "PART_DEMO"
    assert "PART_DEMO" in app._active.Layers.items


async def test_handle_roundtrip_get_and_erase(backend_factory):
    app = FakeApplication()
    backend = backend_factory(app)
    await backend.initialize()

    created = await backend.create_line(0, 0, 10, 0)
    handle = created.payload["handle"]

    fetched = await backend.entity_get(handle)
    assert fetched.ok
    assert fetched.payload["type"] == "LINE"
    assert fetched.payload["start"] == [0.0, 0.0]

    erased = await backend.entity_erase(handle)
    assert erased.ok
    assert app._active.ModelSpace.entities[0].deleted is True


async def test_unknown_handle_returns_friendly_error(backend_factory):
    app = FakeApplication()
    backend = backend_factory(app)
    await backend.initialize()

    result = await backend.entity_get("DEADBEEF")
    assert result.ok is False
    assert "not found" in result.error


async def test_entity_list_and_count(backend_factory):
    app = FakeApplication()
    backend = backend_factory(app)
    await backend.initialize()

    await backend.create_line(0, 0, 1, 0, layer="A")
    await backend.create_circle(0, 0, 1, layer="B")

    listing = await backend.entity_list()
    assert listing.payload["count"] == 2
    assert {e["layer"] for e in listing.payload["entities"]} == {"A", "B"}

    only_a = await backend.entity_list(layer="A")
    assert only_a.payload["count"] == 1
    assert (await backend.entity_count(layer="A")).payload["count"] == 1


# ---------------------------------------------------------------------------
# Dimensions
# ---------------------------------------------------------------------------


async def test_linear_dimension_uses_autocad_orientation(backend_factory):
    app = FakeApplication()
    backend = backend_factory(app)
    await backend.initialize()

    result = await backend.create_dimension_linear(0, 0, 100, 0, 50, -15)
    assert result.ok

    name, args = app._active.ModelSpace.calls[-1]
    assert name == "AddDimRotated"
    assert args[0] == (0.0, 0.0, 0.0)
    assert args[1] == (100.0, 0.0, 0.0)
    assert args[3] == pytest.approx(0.0)


async def test_vertical_linear_dimension_uses_90_degrees(backend_factory):
    app = FakeApplication()
    backend = backend_factory(app)
    await backend.initialize()

    result = await backend.create_dimension_linear(100, 0, 100, 60, 115, 30)
    assert result.ok

    name, args = app._active.ModelSpace.calls[-1]
    assert name == "AddDimRotated"
    assert args[3] == pytest.approx(math.pi / 2)


async def test_linear_dimension_falls_back_to_sendcommand(backend_factory):
    app = FakeApplication()
    backend = backend_factory(app)
    await backend.initialize()

    model_space = app._active.ModelSpace

    def boom(*args, **kwargs):
        model_space._record("AddDimRotated")
        raise FakeComError(0x80020009)

    model_space.AddDimRotated = boom

    result = await backend.create_dimension_linear(0, 0, 100, 0, 50, -15)
    assert result.ok
    assert result.payload["via"] == "sendcommand"
    assert app._active.commands
    assert "DIMLINEAR" in app._active.commands[-1]


async def test_diameter_dimension_places_opposite_chord_points(backend_factory):
    app = FakeApplication()
    backend = backend_factory(app)
    await backend.initialize()

    result = await backend.create_dimension_diameter(50, 30, 10, 0)
    assert result.ok

    name, args = app._active.ModelSpace.calls[-1]
    assert name == "AddDimDiametric"
    assert args[0] == pytest.approx((60.0, 30.0, 0.0))
    assert args[1] == pytest.approx((40.0, 30.0, 0.0))


async def test_aligned_dimension_offsets_along_normal(backend_factory):
    app = FakeApplication()
    backend = backend_factory(app)
    await backend.initialize()

    result = await backend.create_dimension_aligned(100, 0, 100, 60, 15)
    assert result.ok

    name, args = app._active.ModelSpace.calls[-1]
    assert name == "AddDimAligned"
    # Vertical line: normal points +X, so the location sits 15 units right.
    assert args[2][0] == pytest.approx(115.0)
    assert args[2][1] == pytest.approx(30.0)


# ---------------------------------------------------------------------------
# Retry / timeout / unsupported
# ---------------------------------------------------------------------------


async def test_busy_hresult_is_retried(backend_factory):
    app = FakeApplication()
    backend = backend_factory(app)
    await backend.initialize()

    app._active.ModelSpace.fail_active_calls = 2
    app._active.ModelSpace.fail_hresult = 0x80010001

    result = await backend.create_line(0, 0, 1, 1)
    assert result.ok
    assert backend._retry_count == 2


async def test_slow_call_times_out_without_blocking(backend_factory):
    app = FakeApplication()
    backend = backend_factory(app, call_timeout=0.05)
    await backend.initialize()

    app._active.ModelSpace.delay = 0.5
    started = time.monotonic()
    result = await backend.create_line(0, 0, 1, 1)
    elapsed = time.monotonic() - started

    assert result.ok is False
    assert "did not return" in result.error
    assert elapsed < 0.4


async def test_fillet_and_chamfer_are_explicitly_unsupported(backend_factory):
    app = FakeApplication()
    backend = backend_factory(app)
    await backend.initialize()

    fillet = await backend.entity_fillet("A", "B", 5)
    chamfer = await backend.entity_chamfer("A", "B", 2, 2)

    assert fillet.ok is False and "ActiveX" in fillet.error
    assert chamfer.ok is False and "ActiveX" in chamfer.error


# ---------------------------------------------------------------------------
# Layers, blocks, view
# ---------------------------------------------------------------------------


async def test_layer_create_color_and_freeze(backend_factory):
    app = FakeApplication()
    backend = backend_factory(app)
    await backend.initialize()

    created = await backend.layer_create("PART_DEMO", "red")
    assert created.ok
    assert created.payload["color"] == 1

    await backend.layer_freeze("PART_DEMO")
    assert app._active.Layers.items["PART_DEMO"].Freeze is True

    await backend.layer_set_current("PART_DEMO")
    assert app._active.ActiveLayer.Name == "PART_DEMO"

    listing = await backend.layer_list()
    names = {layer["name"] for layer in listing.payload["layers"]}
    assert "PART_DEMO" in names


async def test_block_define_and_insert(backend_factory):
    app = FakeApplication()
    backend = backend_factory(app)
    await backend.initialize()

    defined = await backend.block_define(
        "TAG", [{"type": "line", "x1": 0, "y1": 0, "x2": 1, "y2": 1}]
    )
    assert defined.ok
    assert app._active.Blocks.items["TAG"].entities

    inserted = await backend.block_insert("TAG", 5, 5, 2, 90)
    assert inserted.ok
    assert inserted.payload["block"] == "TAG"


async def test_zoom_and_plot(backend_factory):
    app = FakeApplication()
    backend = backend_factory(app)
    await backend.initialize()

    assert (await backend.zoom_extents()).ok
    assert app.zoom_calls[-1] == "extents"

    assert (await backend.zoom_window(0, 0, 100, 60)).ok
    assert app.zoom_calls[-1][0] == "window"

    plotted = await backend.drawing_plot_pdf("C:/tmp/out.pdf")
    assert plotted.ok
    assert app._active.Plot.plot_file == "C:/tmp/out.pdf"


async def test_status_after_close_is_disconnected(backend_factory):
    app = FakeApplication()
    backend = backend_factory(app)
    await backend.initialize()

    await backend.close()
    status = await backend.status()

    assert status.ok
    assert status.payload["connected"] is False
