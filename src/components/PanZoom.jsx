import { useState, useEffect, useRef, useImperativeHandle, forwardRef } from "react";

const ZOOM_FACTOR = 1.15;
const WHEEL_FACTOR = 1.1;
const MIN_ZOOM = 0.1;
const MAX_ZOOM = 10;

/**
 * A container that adds mouse-drag panning and scroll-wheel zooming to its children.
 *
 * Imperative handle (via ref):
 *   ref.current.zoomIn()
 *   ref.current.zoomOut()
 *   ref.current.reset()    ← resets both zoom and pan to initial state
 */
const PanZoom = forwardRef(function PanZoom({ children }, ref) {
  const [zoom, setZoom]       = useState(1);
  const [pan, setPan]         = useState({ x: 0, y: 0 });
  const [isPanning, setIsPanning] = useState(false);

  // Refs keep event-handler callbacks stable across renders
  const zoomRef            = useRef(1);
  const panRef             = useRef({ x: 0, y: 0 });
  const isPanningRef       = useRef(false);
  const dragStartRef       = useRef({ x: 0, y: 0 });
  const panAtDragStartRef  = useRef({ x: 0, y: 0 });
  const containerRef       = useRef(null);

  const applyZoom = (next) => {
    const clamped = Math.min(MAX_ZOOM, Math.max(MIN_ZOOM, next));
    zoomRef.current = clamped;
    setZoom(clamped);
  };

  useImperativeHandle(ref, () => ({
    zoomIn:  () => applyZoom(zoomRef.current * ZOOM_FACTOR),
    zoomOut: () => applyZoom(zoomRef.current / ZOOM_FACTOR),
    reset:   () => {
      applyZoom(1);
      panRef.current = { x: 0, y: 0 };
      setPan({ x: 0, y: 0 });
    },
  }));

  // ── Mouse pan ─────────────────────────────────────────────────────────────

  const handleMouseDown = (e) => {
    if (e.button !== 0) return;
    e.preventDefault();
    isPanningRef.current = true;
    setIsPanning(true);
    dragStartRef.current      = { x: e.clientX, y: e.clientY };
    panAtDragStartRef.current = { ...panRef.current };
  };

  // Listen on window so dragging continues even when the cursor leaves the element
  useEffect(() => {
    const onMove = (e) => {
      if (!isPanningRef.current) return;
      const next = {
        x: panAtDragStartRef.current.x + (e.clientX - dragStartRef.current.x),
        y: panAtDragStartRef.current.y + (e.clientY - dragStartRef.current.y),
      };
      panRef.current = next;
      setPan(next);
    };
    const onUp = () => {
      if (!isPanningRef.current) return;
      isPanningRef.current = false;
      setIsPanning(false);
    };
    window.addEventListener("mousemove", onMove);
    window.addEventListener("mouseup", onUp);
    return () => {
      window.removeEventListener("mousemove", onMove);
      window.removeEventListener("mouseup", onUp);
    };
  }, []);

  // ── Scroll-wheel zoom ─────────────────────────────────────────────────────
  // Must be non-passive so we can preventDefault() and stop page scroll.

  useEffect(() => {
    const el = containerRef.current;
    if (!el) return;
    const onWheel = (e) => {
      e.preventDefault();
      const factor = e.deltaY < 0 ? WHEEL_FACTOR : 1 / WHEEL_FACTOR;
      applyZoom(zoomRef.current * factor);
    };
    el.addEventListener("wheel", onWheel, { passive: false });
    return () => el.removeEventListener("wheel", onWheel);
  }, []);

  return (
    <div
      ref={containerRef}
      className="graph-wrapper"
      onMouseDown={handleMouseDown}
      style={{ cursor: isPanning ? "grabbing" : "grab" }}
    >
      <div
        style={{
          transform: `translate(${pan.x}px, ${pan.y}px) scale(${zoom})`,
          transformOrigin: "top left",
        }}
      >
        {children}
      </div>
    </div>
  );
});

export default PanZoom;
