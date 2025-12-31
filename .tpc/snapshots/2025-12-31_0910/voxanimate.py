"""
voxanimate.py
Animation preservation system for Voxsmith.

Captures and restores PowerPoint animation timelines when audio is inserted,
working around COM API's inherent animation scrambling behavior.
"""

def snapshot_slide_animations(slide) -> dict:
    """
    Capture complete animation state before audio insertion.
    
    IMPORTANT: Excludes media/audio effects - we only want shape animations.
    
    Returns:
        {
            "effects": [...],
            "has_text_animations": True/False,  # NEW: Flag for text animations
            "text_animation_shapes": [...]      # NEW: Names of affected shapes
        }
    """
    try:
        seq = slide.TimeLine.MainSequence
        snapshot = {
            "effects": [],
            "has_text_animations": False,
            "text_animation_shapes": []
        }
        
        for i in range(1, seq.Count + 1):
            effect = seq.Item(i)
            
            # CRITICAL FIX: Skip media/audio effects
            # We only want to preserve shape animations, not audio
            try:
                if int(effect.EffectType) == 83:  # msoAnimEffectMediaPlay
                    continue  # Skip audio effects
            except:
                pass
            
            eff_data = {
                "index": i,
                "shape_id": None,
                "shape_name": None,
                "effect_type": int(effect.EffectType),
                "trigger_type": int(effect.Timing.TriggerType),
                "trigger_delay": float(effect.Timing.TriggerDelayTime),
                "duration": float(effect.Timing.Duration),
                "speed": 1.0,
                "rewind": False,
                "repeat_count": 1,
                "auto_reverse": False,
                "effect_options": {},
                "behaviors": [],
                "text_unit_effect": None,  # NEW: For text animations
                "paragraph": None,         # NEW: Which paragraph (-1 = all, 0+ = specific)
                "text_range_start": None,  # NEW: Character start position
                "text_range_length": None  # NEW: Character length
            }
            
            # Capture shape info if available
            try:
                if effect.Shape:
                    eff_data["shape_id"] = effect.Shape.Id
                    eff_data["shape_name"] = effect.Shape.Name
            except:
                pass
            
            # NEW: Capture text animation properties (for by-paragraph/by-word animations)
            try:
                # TextUnitEffect property (by paragraph, by word, by letter)
                if hasattr(effect, 'TextUnitEffect'):
                    eff_data["text_unit_effect"] = int(effect.TextUnitEffect)
                    # Mark that this slide has text animations
                    snapshot["has_text_animations"] = True
                    if eff_data.get("shape_name"):
                        snapshot["text_animation_shapes"].append(eff_data["shape_name"])
            except:
                pass
            
            try:
                # Paragraph index (-1 = all, 0+ = specific paragraph)
                if hasattr(effect, 'Paragraph'):
                    eff_data["paragraph"] = int(effect.Paragraph)
                    # Mark that this slide has text animations
                    snapshot["has_text_animations"] = True
                    if eff_data.get("shape_name") and eff_data["shape_name"] not in snapshot["text_animation_shapes"]:
                        snapshot["text_animation_shapes"].append(eff_data["shape_name"])
            except:
                pass
            
            try:
                # Text range for character-level targeting
                if hasattr(effect, 'TextRangeStart'):
                    eff_data["text_range_start"] = int(effect.TextRangeStart)
            except:
                pass
            
            try:
                if hasattr(effect, 'TextRangeLength'):
                    eff_data["text_range_length"] = int(effect.TextRangeLength)
            except:
                pass
            
            # Capture additional timing properties safely
            try:
                eff_data["speed"] = float(effect.Timing.Speed)
            except:
                pass
            
            try:
                eff_data["rewind"] = bool(effect.Timing.RewindWhenDone)
            except:
                pass
            
            try:
                eff_data["repeat_count"] = int(effect.Timing.RepeatCount)
            except:
                pass
            
            try:
                eff_data["auto_reverse"] = bool(effect.Timing.AutoReverse)
            except:
                pass
            
            # NEW: Capture effect options (direction, amount, etc.)
            try:
                if hasattr(effect, 'EffectParameters'):
                    params = effect.EffectParameters
                    
                    # Direction (for Wipe, Fly In, etc.)
                    try:
                        eff_data["effect_options"]["direction"] = int(params.Direction)
                    except:
                        pass
                    
                    # Amount (for Grow/Shrink, etc.)
                    try:
                        eff_data["effect_options"]["amount"] = float(params.Amount)
                    except:
                        pass
                    
                    # Font settings (for text effects)
                    try:
                        eff_data["effect_options"]["font_bold"] = bool(params.FontBold)
                    except:
                        pass
                    
                    try:
                        eff_data["effect_options"]["font_italic"] = bool(params.FontItalic)
                    except:
                        pass
                    
                    try:
                        eff_data["effect_options"]["font_size"] = float(params.FontSize)
                    except:
                        pass
                    
                    try:
                        eff_data["effect_options"]["font_underline"] = bool(params.FontUnderline)
                    except:
                        pass
                    
                    # Color settings
                    try:
                        eff_data["effect_options"]["color_rgb"] = int(params.Color.RGB)
                    except:
                        pass
                    
                    try:
                        eff_data["effect_options"]["color2_rgb"] = int(params.Color2.RGB)
                    except:
                        pass
                    
                    # Relative position
                    try:
                        eff_data["effect_options"]["relative"] = bool(params.Relative)
                    except:
                        pass
            except:
                pass
            
            # NEW: Capture behavior properties (smooth start/end, etc.)
            try:
                if hasattr(effect, 'Behaviors'):
                    for j in range(1, effect.Behaviors.Count + 1):
                        behavior = effect.Behaviors.Item(j)
                        behavior_data = {
                            "type": int(behavior.Type) if hasattr(behavior, 'Type') else None
                        }
                        
                        # Timing properties
                        try:
                            behavior_data["accumulate"] = int(behavior.Accumulate)
                        except:
                            pass
                        
                        try:
                            behavior_data["additive"] = int(behavior.Additive)
                        except:
                            pass
                        
                        # Motion behavior properties
                        try:
                            if behavior.Type == 1:  # msoAnimTypeMotion
                                behavior_data["x"] = float(behavior.MotionEffect.FromX)
                                behavior_data["y"] = float(behavior.MotionEffect.FromY)
                                behavior_data["to_x"] = float(behavior.MotionEffect.ToX)
                                behavior_data["to_y"] = float(behavior.MotionEffect.ToY)
                        except:
                            pass
                        
                        # Property effect (for most animations)
                        try:
                            if behavior.Type == 4:  # msoAnimTypeProperty
                                behavior_data["property"] = int(behavior.PropertyEffect.Property)
                                try:
                                    behavior_data["from_value"] = str(behavior.PropertyEffect.From)
                                except:
                                    pass
                                try:
                                    behavior_data["to_value"] = str(behavior.PropertyEffect.To)
                                except:
                                    pass
                        except:
                            pass
                        
                        # Timing behavior
                        try:
                            timing = behavior.Timing
                            behavior_data["smooth_start"] = float(timing.SmoothStart)
                            behavior_data["smooth_end"] = float(timing.SmoothEnd)
                        except:
                            pass
                        
                        eff_data["behaviors"].append(behavior_data)
            except:
                pass
            
            snapshot["effects"].append(eff_data)
        
        return snapshot
    except Exception as e:
        return {"effects": [], "error": str(e)}


def restore_slide_animations(slide, snapshot: dict, audio_shape) -> bool:
    """
    Restore animations from snapshot, with audio at position 1.
    
    Args:
        slide: PowerPoint slide object
        snapshot: Animation state from snapshot_slide_animations()
        audio_shape: The audio shape we just inserted
    
    Returns:
        True if restoration succeeded, False otherwise
    """
    try:
        seq = slide.TimeLine.MainSequence
        
        # STEP 1: Clear ALL effects (including any old audio)
        while seq.Count > 0:
            try:
                seq.Item(1).Delete()
            except:
                break
        
        # STEP 2: Insert audio effect FIRST (position 1)
        msoAnimEffectMediaPlay = 83
        msoAnimTriggerAfterPrevious = 3
        audio_eff = seq.AddEffect(audio_shape, msoAnimEffectMediaPlay)
        audio_eff.Timing.TriggerType = msoAnimTriggerAfterPrevious
        audio_eff.Timing.TriggerDelayTime = 0.0
        
        # STEP 3: Restore original shape animations (audio effects were excluded from snapshot)
        # Track which effects we skip due to text animation complexity
        skipped_text_effects = []
        
        for eff_data in snapshot.get("effects", []):
            # CRITICAL: Skip text animations (by paragraph/word/letter)
            # These have complex internal structure that can't be simply restored
            # Restoring them causes shape duplication and other issues
            if eff_data.get("text_unit_effect") is not None or eff_data.get("paragraph") is not None:
                shape_name = eff_data.get("shape_name", "unknown")
                skipped_text_effects.append(shape_name)
                continue
            
            # Find the shape by ID or name
            target_shape = None
            shape_id = eff_data.get("shape_id")
            shape_name = eff_data.get("shape_name")
            
            if shape_id:
                for shape in slide.Shapes:
                    try:
                        if shape.Id == shape_id:
                            target_shape = shape
                            break
                    except:
                        continue
            
            if not target_shape and shape_name:
                for shape in slide.Shapes:
                    try:
                        if shape.Name == shape_name:
                            target_shape = shape
                            break
                    except:
                        continue
            
            if not target_shape:
                # Shape doesn't exist anymore, skip this effect
                continue
            
            # Recreate the effect
            try:
                new_eff = seq.AddEffect(target_shape, eff_data["effect_type"])
                new_eff.Timing.TriggerType = eff_data["trigger_type"]
                new_eff.Timing.TriggerDelayTime = eff_data["trigger_delay"]
                new_eff.Timing.Duration = eff_data["duration"]
                
                # Restore timing properties
                try:
                    new_eff.Timing.Speed = eff_data.get("speed", 1.0)
                except:
                    pass
                
                try:
                    new_eff.Timing.RewindWhenDone = eff_data.get("rewind", False)
                except:
                    pass
                
                try:
                    new_eff.Timing.RepeatCount = eff_data.get("repeat_count", 1)
                except:
                    pass
                
                try:
                    new_eff.Timing.AutoReverse = eff_data.get("auto_reverse", False)
                except:
                    pass
                
                # NEW: Restore effect options (direction, amount, etc.)
                effect_options = eff_data.get("effect_options", {})
                if effect_options and hasattr(new_eff, 'EffectParameters'):
                    params = new_eff.EffectParameters
                    
                    # Direction
                    if "direction" in effect_options:
                        try:
                            params.Direction = effect_options["direction"]
                        except:
                            pass
                    
                    # Amount
                    if "amount" in effect_options:
                        try:
                            params.Amount = effect_options["amount"]
                        except:
                            pass
                    
                    # Font properties
                    if "font_bold" in effect_options:
                        try:
                            params.FontBold = effect_options["font_bold"]
                        except:
                            pass
                    
                    if "font_italic" in effect_options:
                        try:
                            params.FontItalic = effect_options["font_italic"]
                        except:
                            pass
                    
                    if "font_size" in effect_options:
                        try:
                            params.FontSize = effect_options["font_size"]
                        except:
                            pass
                    
                    if "font_underline" in effect_options:
                        try:
                            params.FontUnderline = effect_options["font_underline"]
                        except:
                            pass
                    
                    # Color settings
                    if "color_rgb" in effect_options:
                        try:
                            params.Color.RGB = effect_options["color_rgb"]
                        except:
                            pass
                    
                    if "color2_rgb" in effect_options:
                        try:
                            params.Color2.RGB = effect_options["color2_rgb"]
                        except:
                            pass
                    
                    # Relative positioning
                    if "relative" in effect_options:
                        try:
                            params.Relative = effect_options["relative"]
                        except:
                            pass
                
                # NEW: Restore behavior properties (smooth start/end, etc.)
                behaviors_data = eff_data.get("behaviors", [])
                if behaviors_data and hasattr(new_eff, 'Behaviors'):
                    # Match behaviors by index (they should correspond)
                    for idx, behavior_data in enumerate(behaviors_data):
                        try:
                            if idx + 1 <= new_eff.Behaviors.Count:
                                behavior = new_eff.Behaviors.Item(idx + 1)
                                
                                # Restore timing properties
                                if "accumulate" in behavior_data:
                                    try:
                                        behavior.Accumulate = behavior_data["accumulate"]
                                    except:
                                        pass
                                
                                if "additive" in behavior_data:
                                    try:
                                        behavior.Additive = behavior_data["additive"]
                                    except:
                                        pass
                                
                                # Restore motion properties
                                if behavior_data.get("type") == 1:  # msoAnimTypeMotion
                                    try:
                                        if "x" in behavior_data:
                                            behavior.MotionEffect.FromX = behavior_data["x"]
                                        if "y" in behavior_data:
                                            behavior.MotionEffect.FromY = behavior_data["y"]
                                        if "to_x" in behavior_data:
                                            behavior.MotionEffect.ToX = behavior_data["to_x"]
                                        if "to_y" in behavior_data:
                                            behavior.MotionEffect.ToY = behavior_data["to_y"]
                                    except:
                                        pass
                                
                                # Restore property effect values
                                if behavior_data.get("type") == 4:  # msoAnimTypeProperty
                                    try:
                                        if "from_value" in behavior_data:
                                            behavior.PropertyEffect.From = behavior_data["from_value"]
                                        if "to_value" in behavior_data:
                                            behavior.PropertyEffect.To = behavior_data["to_value"]
                                    except:
                                        pass
                                
                                # Restore smooth start/end
                                if "smooth_start" in behavior_data or "smooth_end" in behavior_data:
                                    try:
                                        timing = behavior.Timing
                                        if "smooth_start" in behavior_data:
                                            timing.SmoothStart = behavior_data["smooth_start"]
                                        if "smooth_end" in behavior_data:
                                            timing.SmoothEnd = behavior_data["smooth_end"]
                                    except:
                                        pass
                        except:
                            pass
                    
            except Exception as e:
                # This specific effect failed to restore, log but continue
                print(f"Warning: Could not restore effect for {eff_data.get('shape_name', 'unknown')}: {e}")
                continue
        
        # Report skipped text animations (if any)
        if skipped_text_effects:
            unique_shapes = list(set(skipped_text_effects))
            print(f"Note: Skipped {len(skipped_text_effects)} text animation(s) on shape(s): {', '.join(unique_shapes[:3])}")
            print("      Text animations (by paragraph/word) must be manually reapplied after audio insertion.")
        
        return True
        
    except Exception as e:
        print(f"Restore failed: {e}")
        return False


def cleanup_orphaned_audio_effects(slide):
    """
    Remove animation effects for audio shapes that no longer exist.
    Call this before snapshot to ensure clean state.
    """
    try:
        seq = slide.TimeLine.MainSequence
        for i in range(seq.Count, 0, -1):
            try:
                eff = seq.Item(i)
                shape = eff.Shape
                
                # Check if this is a media effect
                if eff.EffectType == 83:  # msoAnimEffectMediaPlay
                    # Try to access the shape
                    try:
                        _ = shape.Name
                    except:
                        # Shape doesn't exist, delete the effect
                        eff.Delete()
                        
            except Exception:
                # Effect or shape reference is broken, try to delete
                try:
                    seq.Item(i).Delete()
                except:
                    pass
    except Exception:
        pass


def should_skip_audio_attachment(snapshot: dict) -> tuple:
    """
    Check if audio attachment should be skipped due to text animations.
    
    Returns:
        (should_skip: bool, reason: str)
    """
    if snapshot.get("has_text_animations", False):
        shapes = snapshot.get("text_animation_shapes", [])
        unique_shapes = list(set(shapes))
        shape_list = ", ".join(unique_shapes[:3])
        if len(unique_shapes) > 3:
            shape_list += f", and {len(unique_shapes) - 3} more"
        
        reason = f"Text animations detected on: {shape_list}"
        return (True, reason)
    
    return (False, "")
