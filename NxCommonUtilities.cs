using System;
using System.Collections.Generic;
using System.Linq;
using NXOpen;
using NXOpen.Features;
using NXOpen.GeometricUtilities;
using NXOpen.UF;
using NXOpen.BlockStyler;
using NXOpen.Utilities;
// using System.Windows.Forms; // Uncomment only if you really need WinForms

namespace NxCommonUtilities
{
    /// <summary>
    /// Common helper utilities for Siemens NXOpen operations:
    /// - Curve & Sketch operations
    /// - Extrude & Boolean body operations
    /// - Offsets, trim, projection & datum utilities
    /// - Measurement & evaluation helpers
    /// 
    /// All methods are defensive:
    /// - Null checks
    /// - Try/catch with logging to NX Listing Window
    /// </summary>
    public static class NxCommonUtility
    {
        private static readonly Session theSession = Session.GetSession();
        private static readonly UFSession theUFSession = UFSession.GetUFSession();
        private static Part workPart => theSession.Parts.Work;

        #region === Logging Helpers ===

        private static void LogInfo(string message)
        {
            try
            {
                theSession.ListingWindow.Open();
                theSession.ListingWindow.WriteLine("[INFO] " + message);
            }
            catch
            {
                // Swallow logging failures – never crash on logging
            }
        }

        private static void LogError(string context, Exception ex)
        {
            try
            {
                theSession.ListingWindow.Open();
                theSession.ListingWindow.WriteLine($"[ERROR] {context}: {ex.Message}");
            }
            catch
            {
                // Swallow logging failures
            }
        }

        #endregion

        #region === Color Utilities ===

        /// <summary>
        /// Change color of all curves in a collection.
        /// </summary>
        public static void ColorChange(List<Curve> curveCollection, int color)
        {
            try
            {
                if (curveCollection == null || curveCollection.Count == 0)
                    return;

                foreach (Curve curve in curveCollection)
                {
                    if (curve != null)
                    {
                        curve.Color = color;
                    }
                }
            }
            catch (Exception ex)
            {
                LogError(nameof(ColorChange), ex);
                throw;
            }
        }

        #endregion

        #region === Extrude Utilities ===

        /// <summary>
        /// Extrudes a set of curves with tangent chaining to create a solid body.
        /// </summary>
        public static Body ExtrudeCurves(Curve[] curves, double distance, Direction direction)
        {
            if (curves == null || curves.Length == 0)
                throw new ArgumentException("Input curves cannot be null or empty.", nameof(curves));

            if (direction == null)
                throw new ArgumentNullException(nameof(direction));

            ExtrudeBuilder extrudeBuilder = null;

            try
            {
                Part wp = workPart;
                extrudeBuilder = wp.Features.CreateExtrudeBuilder(null);

                Section section = wp.Sections.CreateSection(0.0001, 0.001, 0.5);

                foreach (Curve curve in curves)
                {
                    if (curve == null) continue;

                    double angleTolerance = 0.5;
                    double gapTolerance = 0.01;

                    SelectionIntentRule[] tangentRules =
                    {
                        wp.ScRuleFactory.CreateRuleCurveTangent(
                            curve,
                            null,
                            true,
                            angleTolerance,
                            gapTolerance)
                    };

                    section.AddToSection(
                        tangentRules,
                        curve,
                        null,
                        null,
                        new Point3d(0, 0, 0),
                        Section.Mode.Create,
                        false);
                }

                extrudeBuilder.Section = section;
                extrudeBuilder.Direction = direction;
                extrudeBuilder.FeatureOptions.BodyType = FeatureOptions.BodyStyle.Solid;
                extrudeBuilder.Limits.StartExtend.Value.RightHandSide = "-500";
                extrudeBuilder.Limits.EndExtend.Value.RightHandSide = distance.ToString();

                NXObject extrude = extrudeBuilder.Commit();

                Body resultingBody = null;

                if (extrude is Feature feature)
                {
                    foreach (NXObject entity in feature.GetEntities())
                    {
                        if (entity is Body body)
                        {
                            resultingBody = body;
                            break;
                        }
                    }
                }

                if (resultingBody == null)
                    throw new InvalidOperationException("Failed to create an extruded body. Please check the input curves and direction.");

                LogInfo("ExtrudeCurves successful.");
                return resultingBody;
            }
            catch (Exception ex)
            {
                LogError(nameof(ExtrudeCurves), ex);
                throw;
            }
            finally
            {
                extrudeBuilder?.Destroy();
            }
        }

        /// <summary>
        /// Extrude tangent curves and edges as a solid body.
        /// </summary>
        public static Body ExtrudeTangentCurvesAndEdges(List<Curve> curvesAndEdges, double distance, Direction direction)
        {
            Session.UndoMarkId markId1 = theSession.SetUndoMark(Session.MarkVisibility.Visible, "Extrude Tangent Curves");
            ExtrudeBuilder extrudeBuilder = null;

            try
            {
                if (curvesAndEdges == null || curvesAndEdges.Count == 0)
                    throw new ArgumentException("curvesAndEdges cannot be null or empty.", nameof(curvesAndEdges));

                if (direction == null)
                    throw new ArgumentNullException(nameof(direction));

                extrudeBuilder = workPart.Features.CreateExtrudeBuilder(null);
                Section section = workPart.Sections.CreateSection(0.0001, 0.001, 0.5);

                List<SelectionIntentRule> selectionRules = new List<SelectionIntentRule>();
                SelectionIntentRuleOptions ruleOptions = workPart.ScRuleFactory.CreateRuleOptions();
                ruleOptions.SetSelectedFromInactive(false);

                foreach (NXObject obj in curvesAndEdges)
                {
                    if (obj is Curve curve)
                    {
                        SelectionIntentRule tangentRule = workPart.ScRuleFactory.CreateRuleCurveTangent(
                            curve,
                            null,
                            true,
                            0.5,
                            0.01);
                        selectionRules.Add(tangentRule);
                    }
                    else if (obj is Edge edge)
                    {
                        SelectionIntentRule rule = workPart.ScRuleFactory.CreateRuleEdgeTangent(
                            edge,
                            null,
                            true,
                            0.5,
                            false,
                            false,
                            ruleOptions);
                        selectionRules.Add(rule);
                    }
                }

                section.AddToSection(
                    selectionRules.ToArray(),
                    curvesAndEdges[0],
                    null,
                    null,
                    new Point3d(0, 0, 0),
                    Section.Mode.Create,
                    false);

                extrudeBuilder.Section = section;
                extrudeBuilder.Direction = direction;
                extrudeBuilder.FeatureOptions.BodyType = FeatureOptions.BodyStyle.Solid;
                extrudeBuilder.Limits.StartExtend.Value.RightHandSide = "0";
                extrudeBuilder.Limits.EndExtend.Value.RightHandSide = distance.ToString();

                Body resultBody = null;

                Feature extrudeFeature = extrudeBuilder.CommitFeature();
                resultBody = extrudeFeature.GetBodies().FirstOrDefault();

                LogInfo("ExtrudeTangentCurvesAndEdges successful.");
                return resultBody;
            }
            catch (Exception ex)
            {
                LogError(nameof(ExtrudeTangentCurvesAndEdges), ex);
                return null;
            }
            finally
            {
                extrudeBuilder?.Destroy();
                theSession.DeleteUndoMark(markId1, null);
            }
        }

        /// <summary>
        /// Extrude connected curves and edges as a solid.
        /// </summary>
        public static Body ExtrudeConnectedCurvesAndEdges(Curve[] curvesAndEdges, double distance, Direction direction)
        {
            Session.UndoMarkId markId1 = theSession.SetUndoMark(Session.MarkVisibility.Visible, "Extrude Connected Curves");
            ExtrudeBuilder extrudeBuilder = null;

            try
            {
                if (curvesAndEdges == null || curvesAndEdges.Length == 0)
                    throw new ArgumentException("curvesAndEdges cannot be null or empty.", nameof(curvesAndEdges));

                if (direction == null)
                    throw new ArgumentNullException(nameof(direction));

                extrudeBuilder = workPart.Features.CreateExtrudeBuilder(null);
                Section section = workPart.Sections.CreateSection(0.0001, 0.001, 0.5);

                List<SelectionIntentRule> selectionRules = new List<SelectionIntentRule>();

                foreach (NXObject obj in curvesAndEdges)
                {
                    if (obj is Curve curve)
                    {
                        SelectionIntentRule connectedRule = workPart.ScRuleFactory.CreateRuleCurveChain(
                            curve,
                            null,
                            true,
                            0.01);
                        selectionRules.Add(connectedRule);
                    }
                    else if (obj is Edge edge)
                    {
                        SelectionIntentRule rule = workPart.ScRuleFactory.CreateRuleEdgeChain(
                            edge,
                            null,
                            true);
                        selectionRules.Add(rule);
                    }
                }

                section.AddToSection(
                    selectionRules.ToArray(),
                    curvesAndEdges[0],
                    null,
                    null,
                    new Point3d(0, 0, 0),
                    Section.Mode.Create,
                    false);

                extrudeBuilder.Section = section;
                extrudeBuilder.Direction = direction;
                extrudeBuilder.FeatureOptions.BodyType = FeatureOptions.BodyStyle.Solid;
                extrudeBuilder.Limits.StartExtend.Value.RightHandSide = "-250";
                extrudeBuilder.Limits.EndExtend.Value.RightHandSide = distance.ToString();

                Feature extrudeFeature = extrudeBuilder.CommitFeature();
                Body resultBody = extrudeFeature.GetBodies().FirstOrDefault();

                LogInfo("ExtrudeConnectedCurvesAndEdges successful.");
                return resultBody;
            }
            catch (Exception ex)
            {
                LogError(nameof(ExtrudeConnectedCurvesAndEdges), ex);
                return null;
            }
            finally
            {
                extrudeBuilder?.Destroy();
                theSession.DeleteUndoMark(markId1, null);
            }
        }

        /// <summary>
        /// Extrudes a sketch as a solid body.
        /// </summary>
        public static Body[] ExtrudeSketchAsSolid(Sketch sketch, double distance, Direction direction)
        {
            if (sketch == null)
                throw new ArgumentNullException(nameof(sketch));

            if (direction == null)
                throw new ArgumentNullException(nameof(direction));

            ExtrudeBuilder extrudeBuilder = null;

            try
            {
                Part wp = workPart;
                extrudeBuilder = wp.Features.CreateExtrudeBuilder(null);

                Section section = wp.Sections.CreateSection(0.0001, 0.001, 0.5);

                NXObject[] sketchGeometry = sketch.GetAllGeometry();
                IBaseCurve[] baseCurves = Array.ConvertAll(sketchGeometry, item => item as IBaseCurve);
                baseCurves = Array.FindAll(baseCurves, curve => curve != null);

                if (baseCurves.Length == 0)
                {
                    LogInfo("ExtrudeSketchAsSolid: No valid curves found in the sketch.");
                    return null;
                }

                SelectionIntentRule[] sketchRules =
                {
                    wp.ScRuleFactory.CreateRuleBaseCurveDumb(baseCurves)
                };

                section.AddToSection(
                    sketchRules,
                    sketch,
                    null,
                    null,
                    new Point3d(0, 0, 0),
                    Section.Mode.Create,
                    false);

                extrudeBuilder.Section = section;
                extrudeBuilder.Direction = direction;
                extrudeBuilder.FeatureOptions.BodyType = FeatureOptions.BodyStyle.Solid;
                extrudeBuilder.Limits.StartExtend.Value.RightHandSide = "0";
                extrudeBuilder.Limits.EndExtend.Value.RightHandSide = distance.ToString();

                Feature extrudeFeature = extrudeBuilder.CommitFeature();
                Body[] resultBodies = extrudeFeature.GetBodies();

                LogInfo("ExtrudeSketchAsSolid successful.");
                return resultBodies;
            }
            catch (Exception ex)
            {
                LogError(nameof(ExtrudeSketchAsSolid), ex);
                return null;
            }
            finally
            {
                extrudeBuilder?.Destroy();
            }
        }

        #endregion

        #region === Boolean / Body Operations ===

        /// <summary>
        /// Boolean Subtract with multiple tool bodies.
        /// </summary>
        public static BooleanFeature Subtract(Body target, Body[] tools)
        {
            BooleanBuilder booleanBuilder = null;

            try
            {
                if (target == null)
                    throw new ArgumentNullException(nameof(target));

                if (tools == null || tools.Length == 0)
                    throw new ArgumentException("Tools array cannot be null or empty.", nameof(tools));

                booleanBuilder = workPart.Features.CreateBooleanBuilder(null);
                booleanBuilder.Target = target;

                foreach (Body tool in tools)
                {
                    if (tool != null)
                        booleanBuilder.Tools.Add(tool);
                }

                booleanBuilder.Operation = Feature.BooleanType.Subtract;

                BooleanFeature result = (BooleanFeature)booleanBuilder.CommitFeature();
                LogInfo("Subtract successful.");
                return result;
            }
            catch (Exception ex)
            {
                LogError(nameof(Subtract), ex);
                throw;
            }
            finally
            {
                booleanBuilder?.Destroy();
            }
        }

        /// <summary>
        /// Boolean Subtract with a single tool body and CopyTools flag.
        /// </summary>
        public static BooleanFeature SubtractFunc(Body target, Body tool, bool copyTools)
        {
            BooleanBuilder booleanBuilder = null;

            try
            {
                if (target == null)
                    throw new ArgumentNullException(nameof(target));

                if (tool == null)
                    throw new ArgumentNullException(nameof(tool));

                booleanBuilder = workPart.Features.CreateBooleanBuilder(null);
                booleanBuilder.Target = target;
                booleanBuilder.Tools.Add(tool);
                booleanBuilder.Operation = Feature.BooleanType.Subtract;
                booleanBuilder.CopyTools = copyTools;

                BooleanFeature result = (BooleanFeature)booleanBuilder.CommitFeature();
                LogInfo("SubtractFunc successful.");
                return result;
            }
            catch (Exception ex)
            {
                LogError(nameof(SubtractFunc), ex);
                throw;
            }
            finally
            {
                booleanBuilder?.Destroy();
            }
        }

        #endregion

        #region === Extract Body Utilities ===

        /// <summary>
        /// Extract selected bodies into a single dumb body.
        /// </summary>
        public static Body ExtractBodiesAsSingleBody(Body[] selectedObjects)
        {
            ExtractFaceBuilder extractFaceBuilder = null;

            try
            {
                if (selectedObjects == null || selectedObjects.Length == 0)
                {
                    LogInfo("ExtractBodiesAsSingleBody: No bodies selected for extraction.");
                    return null;
                }

                extractFaceBuilder = workPart.Features.CreateExtractFaceBuilder(null);
                extractFaceBuilder.Type = ExtractFaceBuilder.ExtractType.Body;
                extractFaceBuilder.InheritMaterial = true;
                extractFaceBuilder.Associative = true;
                extractFaceBuilder.HideOriginal = false;

                SelectionIntentRuleOptions selectionIntentRuleOptions = workPart.ScRuleFactory.CreateRuleOptions();
                selectionIntentRuleOptions.SetSelectedFromInactive(false);

                BodyDumbRule bodyDumbRule =
                    workPart.ScRuleFactory.CreateRuleBodyDumb(selectedObjects, true, selectionIntentRuleOptions);

                extractFaceBuilder.ExtractBodyCollector.ReplaceRules(new SelectionIntentRule[] { bodyDumbRule }, false);
                selectionIntentRuleOptions.Dispose();

                NXObject extractedObject = extractFaceBuilder.Commit();
                Body extractedBody = extractedObject as Body;

                if (extractedBody != null)
                {
                    LogInfo("ExtractBodiesAsSingleBody successful: " + extractedBody.Name);
                }
                else
                {
                    LogInfo("ExtractBodiesAsSingleBody: Failed to extract a single body.");
                }

                return extractedBody;
            }
            catch (Exception ex)
            {
                LogError(nameof(ExtractBodiesAsSingleBody), ex);
                return null;
            }
            finally
            {
                extractFaceBuilder?.Destroy();
            }
        }

        /// <summary>
        /// Extract a single body and return the extracted copy.
        /// </summary>
        public static Body[] ExtractBodies(Body selectedBody)
        {
            ExtractFaceBuilder extractFaceBuilder = null;
            List<Body> extractedBodies = new List<Body>();

            try
            {
                if (selectedBody == null)
                {
                    LogInfo("ExtractBodies: Selected object is not a body.");
                    return Array.Empty<Body>();
                }

                extractFaceBuilder = workPart.Features.CreateExtractFaceBuilder(null);
                extractFaceBuilder.Type = ExtractFaceBuilder.ExtractType.Body;
                extractFaceBuilder.InheritMaterial = true;
                extractFaceBuilder.Associative = true;
                extractFaceBuilder.HideOriginal = false;

                SelectionIntentRuleOptions selectionIntentRuleOptions = workPart.ScRuleFactory.CreateRuleOptions();
                selectionIntentRuleOptions.SetSelectedFromInactive(false);

                Body[] bodies = { selectedBody };
                BodyDumbRule bodyDumbRule =
                    workPart.ScRuleFactory.CreateRuleBodyDumb(bodies, true, selectionIntentRuleOptions);

                extractFaceBuilder.ExtractBodyCollector.ReplaceRules(new SelectionIntentRule[] { bodyDumbRule }, false);
                selectionIntentRuleOptions.Dispose();

                NXObject extractedObject = extractFaceBuilder.Commit();

                if (extractedObject is Body extractedBody)
                {
                    extractedBodies.Add(extractedBody);
                    LogInfo("ExtractBodies successful: " + extractedBody.Name);
                }

                return extractedBodies.ToArray();
            }
            catch (Exception ex)
            {
                LogError(nameof(ExtractBodies), ex);
                return Array.Empty<Body>();
            }
            finally
            {
                extractFaceBuilder?.Destroy();
            }
        }

        /// <summary>
        /// Extracts a body and returns the extracted copy. If extraction fails, returns original body.
        /// </summary>
        public static Body ExtractBody(Body selectedBody)
        {
            ExtractFaceBuilder extractFaceBuilder = null;

            try
            {
                if (selectedBody == null)
                {
                    LogInfo("ExtractBody: Selected object is not a body.");
                    return null;
                }

                extractFaceBuilder = workPart.Features.CreateExtractFaceBuilder(null);
                extractFaceBuilder.Type = ExtractFaceBuilder.ExtractType.Body;
                extractFaceBuilder.InheritMaterial = true;
                extractFaceBuilder.Associative = true;
                extractFaceBuilder.HideOriginal = false;

                SelectionIntentRuleOptions selectionIntentRuleOptions = workPart.ScRuleFactory.CreateRuleOptions();
                selectionIntentRuleOptions.SetSelectedFromInactive(false);

                Body[] bodies = { selectedBody };
                BodyDumbRule bodyDumbRule =
                    workPart.ScRuleFactory.CreateRuleBodyDumb(bodies, true, selectionIntentRuleOptions);

                extractFaceBuilder.ExtractBodyCollector.ReplaceRules(new SelectionIntentRule[] { bodyDumbRule }, false);
                selectionIntentRuleOptions.Dispose();

                NXObject extractedObject = extractFaceBuilder.Commit();
                extractFaceBuilder.Destroy();
                extractFaceBuilder = null;

                if (extractedObject is Body extractedBody)
                {
                    LogInfo("ExtractBody successful: " + extractedBody.Name);
                    return extractedBody;
                }

                LogInfo("ExtractBody: Failed to extract, returning original body.");
                return selectedBody;
            }
            catch (Exception ex)
            {
                LogError(nameof(ExtractBody), ex);
                return selectedBody;
            }
            finally
            {
                extractFaceBuilder?.Destroy();
            }
        }

        #endregion

        #region === Trim Body By Volume ===

        /// <summary>
        /// Trim body based on volume percentage (less than 50% cut).
        /// </summary>
        public static Body TrimBodyVolume(string extrudeName, string panelName, bool direction)
        {
            TrimBody2Builder trimBody2Builder = null;
            Plane plane1 = null;
            Expression expression1 = null;
            Expression expression2 = null;

            Session.UndoMarkId markId1 = theSession.SetUndoMark(Session.MarkVisibility.Visible, "TrimBodyVolume Start");

            try
            {
                Body body1 = workPart.Bodies.FindObject(extrudeName) as Body;
                Body body2 = workPart.Bodies.FindObject(panelName) as Body;

                if (body1 == null || body2 == null)
                    throw new ArgumentException("Could not find target or panel body by name.");

                TrimBody2 nullTrimBody = null;
                trimBody2Builder = workPart.Features.CreateTrimBody2Builder(nullTrimBody);

                Point3d origin1 = new Point3d(0.0, 0.0, 0.0);
                Vector3d normal1 = new Vector3d(0.0, 0.0, 1.0);
                plane1 = workPart.Planes.CreatePlane(origin1, normal1, SmartObject.UpdateOption.WithinModeling);

                trimBody2Builder.BooleanTool.FacePlaneTool.ToolPlane = plane1;

                Unit unit1 = workPart.UnitCollection.FindObject("MilliMeter") as Unit;
                expression1 = workPart.Expressions.CreateSystemExpressionWithUnits("0", unit1);
                expression2 = workPart.Expressions.CreateSystemExpressionWithUnits("0", unit1);

                trimBody2Builder.Tolerance = 0.01;
                theSession.SetUndoMarkName(markId1, "Trim Body Dialog");

                trimBody2Builder.BooleanTool.ExtrudeRevolveTool.ToolSection.DistanceTolerance = 0.01;
                trimBody2Builder.BooleanTool.ExtrudeRevolveTool.ToolSection.ChainingTolerance = 0.0095;

                ScCollector scCollector1 = workPart.ScCollectors.CreateCollector();
                SelectionIntentRuleOptions selectionIntentRuleOptions1 = workPart.ScRuleFactory.CreateRuleOptions();
                selectionIntentRuleOptions1.SetSelectedFromInactive(false);

                Body[] bodies1 = { body1 };
                BodyDumbRule bodyDumbRule1 =
                    workPart.ScRuleFactory.CreateRuleBodyDumb(bodies1, true, selectionIntentRuleOptions1);
                selectionIntentRuleOptions1.Dispose();

                SelectionIntentRule[] rules1 = { bodyDumbRule1 };
                scCollector1.ReplaceRules(rules1, false);

                double vol1 = MeasureBodyVolume(body1);
                trimBody2Builder.TargetBodyCollector = scCollector1;

                SelectionIntentRuleOptions selectionIntentRuleOptions2 = workPart.ScRuleFactory.CreateRuleOptions();
                selectionIntentRuleOptions2.SetSelectedFromInactive(false);

                FaceBodyRule faceBodyRule1 =
                    workPart.ScRuleFactory.CreateRuleFaceBody(body2, selectionIntentRuleOptions2);
                selectionIntentRuleOptions2.Dispose();

                SelectionIntentRule[] rules2 = { faceBodyRule1 };
                trimBody2Builder.BooleanTool.FacePlaneTool.ToolFaces.FaceCollector.ReplaceRules(rules2, false);

                trimBody2Builder.BooleanTool.ReverseDirection = direction;

                Session.UndoMarkId markId2 = theSession.SetUndoMark(Session.MarkVisibility.Invisible, "Trim Body");
                NXObject nXObject1 = trimBody2Builder.Commit();

                double vol2 = MeasureBodyVolume(body1);
                double volumeDifference = Math.Abs(vol1 - vol2);
                double volumePercentage = (volumeDifference / vol1) * 100.0;

                if (volumePercentage < 50.0)
                {
                    theSession.UndoToMark(markId2, null);
                    trimBody2Builder.BooleanTool.ReverseDirection = !direction;
                    nXObject1 = trimBody2Builder.Commit();
                    vol2 = MeasureBodyVolume(body1);
                }

                theSession.SetUndoMarkName(markId1, "Trim Body");
                LogInfo("TrimBodyVolume successful.");
                return body1;
            }
            catch (Exception ex)
            {
                LogError(nameof(TrimBodyVolume), ex);
                return null;
            }
            finally
            {
                trimBody2Builder?.Destroy();

                try
                {
                    if (expression2 != null)
                        workPart.Expressions.Delete(expression2);
                }
                catch (NXException ex)
                {
                    ex.AssertErrorCode(1050029);
                }

                try
                {
                    if (expression1 != null)
                        workPart.Expressions.Delete(expression1);
                }
                catch (NXException ex)
                {
                    ex.AssertErrorCode(1050029);
                }

                plane1?.DestroyPlane();
                theSession.DeleteUndoMark(markId1, null);
            }
        }

        /// <summary>
        /// Trim body based on volume percentage (more than 50% cut).
        /// </summary>
        public static Body TrimBodyVolumeL(string extrudeName, string panelName, bool direction)
        {
            TrimBody2Builder trimBody2Builder = null;
            Plane plane1 = null;
            Expression expression1 = null;
            Expression expression2 = null;

            Session.UndoMarkId markId1 = theSession.SetUndoMark(Session.MarkVisibility.Visible, "TrimBodyVolumeL Start");

            try
            {
                Body body1 = workPart.Bodies.FindObject(extrudeName) as Body;
                Body body2 = workPart.Bodies.FindObject(panelName) as Body;

                if (body1 == null || body2 == null)
                    throw new ArgumentException("Could not find target or panel body by name.");

                TrimBody2 nullTrimBody = null;
                trimBody2Builder = workPart.Features.CreateTrimBody2Builder(nullTrimBody);

                Point3d origin1 = new Point3d(0.0, 0.0, 0.0);
                Vector3d normal1 = new Vector3d(0.0, 0.0, 1.0);
                plane1 = workPart.Planes.CreatePlane(origin1, normal1, SmartObject.UpdateOption.WithinModeling);

                trimBody2Builder.BooleanTool.FacePlaneTool.ToolPlane = plane1;

                Unit unit1 = workPart.UnitCollection.FindObject("MilliMeter") as Unit;
                expression1 = workPart.Expressions.CreateSystemExpressionWithUnits("0", unit1);
                expression2 = workPart.Expressions.CreateSystemExpressionWithUnits("0", unit1);

                trimBody2Builder.Tolerance = 0.01;
                theSession.SetUndoMarkName(markId1, "Trim Body Dialog");

                trimBody2Builder.BooleanTool.ExtrudeRevolveTool.ToolSection.DistanceTolerance = 0.01;
                trimBody2Builder.BooleanTool.ExtrudeRevolveTool.ToolSection.ChainingTolerance = 0.0095;

                ScCollector scCollector1 = workPart.ScCollectors.CreateCollector();
                SelectionIntentRuleOptions selectionIntentRuleOptions1 = workPart.ScRuleFactory.CreateRuleOptions();
                selectionIntentRuleOptions1.SetSelectedFromInactive(false);

                Body[] bodies1 = { body1 };
                BodyDumbRule bodyDumbRule1 =
                    workPart.ScRuleFactory.CreateRuleBodyDumb(bodies1, true, selectionIntentRuleOptions1);
                selectionIntentRuleOptions1.Dispose();

                SelectionIntentRule[] rules1 = { bodyDumbRule1 };
                scCollector1.ReplaceRules(rules1, false);

                double vol1 = MeasureBodyVolume(body1);
                trimBody2Builder.TargetBodyCollector = scCollector1;

                SelectionIntentRuleOptions selectionIntentRuleOptions2 = workPart.ScRuleFactory.CreateRuleOptions();
                selectionIntentRuleOptions2.SetSelectedFromInactive(false);

                FaceBodyRule faceBodyRule1 =
                    workPart.ScRuleFactory.CreateRuleFaceBody(body2, selectionIntentRuleOptions2);
                selectionIntentRuleOptions2.Dispose();

                SelectionIntentRule[] rules2 = { faceBodyRule1 };
                trimBody2Builder.BooleanTool.FacePlaneTool.ToolFaces.FaceCollector.ReplaceRules(rules2, false);

                trimBody2Builder.BooleanTool.ReverseDirection = direction;

                Session.UndoMarkId markId2 = theSession.SetUndoMark(Session.MarkVisibility.Invisible, "Trim Body");
                NXObject nXObject1 = trimBody2Builder.Commit();

                double vol2 = MeasureBodyVolume(body1);
                double volumeDifference = Math.Abs(vol1 - vol2);
                double volumePercentage = (volumeDifference / vol1) * 100.0;

                if (volumePercentage > 50.0)
                {
                    theSession.UndoToMark(markId2, null);
                    trimBody2Builder.BooleanTool.ReverseDirection = !direction;
                    nXObject1 = trimBody2Builder.Commit();
                    vol2 = MeasureBodyVolume(body1);
                }

                theSession.SetUndoMarkName(markId1, "Trim Body");
                LogInfo("TrimBodyVolumeL successful.");
                return body1;
            }
            catch (Exception ex)
            {
                LogError(nameof(TrimBodyVolumeL), ex);
                return null;
            }
            finally
            {
                trimBody2Builder?.Destroy();

                try
                {
                    if (expression2 != null)
                        workPart.Expressions.Delete(expression2);
                }
                catch (NXException ex)
                {
                    ex.AssertErrorCode(1050029);
                }

                try
                {
                    if (expression1 != null)
                        workPart.Expressions.Delete(expression1);
                }
                catch (NXException ex)
                {
                    ex.AssertErrorCode(1050029);
                }

                plane1?.DestroyPlane();
                theSession.DeleteUndoMark(markId1, null);
            }
        }

        #endregion

        #region === Measurement & Evaluation ===

        public static double MeasureBodyVolume(Body body)
        {
            try
            {
                if (body == null)
                    throw new ArgumentNullException(nameof(body));

                MeasureBodies measureBodies =
                    workPart.MeasureManager.NewMassProperties(
                        new Unit[] { workPart.UnitCollection.FindObject("CubicMilliMeter") },
                        0.01,
                        new Body[] { body });

                return measureBodies.Volume;
            }
            catch (Exception ex)
            {
                LogError(nameof(MeasureBodyVolume), ex);
                throw;
            }
        }

        public static Point3d[] AskCurveEnds(Curve theCurve)
        {
            if (theCurve == null)
                throw new ArgumentNullException(nameof(theCurve), "The input curve cannot be null.");

            double[] limits = new double[2];
            IntPtr evaluator = IntPtr.Zero;
            double[] start = new double[3];
            double[] end = new double[3];

            try
            {
                theUFSession.Eval.Initialize2(theCurve.Tag, out evaluator);
                theUFSession.Eval.AskLimits(evaluator, limits);

                if (limits.Length != 2)
                    throw new Exception("Curve parameter limits are not valid.");

                theUFSession.Eval.Evaluate(evaluator, 0, limits[0], start, new double[] { });
                theUFSession.Eval.Evaluate(evaluator, 0, limits[1], end, new double[] { });

                return new[]
                {
                    new Point3d(start[0], start[1], start[2]),
                    new Point3d(end[0], end[1], end[2])
                };
            }
            catch (NXException ex)
            {
                LogError(nameof(AskCurveEnds), ex);
                throw new Exception($"Error while evaluating curve endpoints: {ex.Message}", ex);
            }
            finally
            {
                if (evaluator != IntPtr.Zero)
                {
                    theUFSession.Eval.Free(evaluator);
                }
            }
        }

        public static List<Point3d> GetUniqueCurveEndpoints(List<Curve[]> curveArrays, double tolerance)
        {
            try
            {
                if (curveArrays == null || curveArrays.Count == 0)
                    throw new ArgumentNullException(nameof(curveArrays), "The input curve arrays list cannot be null or empty.");

                Dictionary<Point3d, int> endpointCounts = new Dictionary<Point3d, int>(new Point3dComparer(tolerance));

                foreach (Curve[] curves in curveArrays)
                {
                    if (curves == null || curves.Length == 0)
                        throw new ArgumentNullException(nameof(curves), "One of the curve arrays is null or empty.");

                    foreach (Curve curve in curves)
                    {
                        double[] limits = new double[2];
                        IntPtr evaluator = IntPtr.Zero;
                        double[] start = new double[3];
                        double[] end = new double[3];

                        try
                        {
                            theUFSession.Eval.Initialize2(curve.Tag, out evaluator);
                            theUFSession.Eval.AskLimits(evaluator, limits);
                            theUFSession.Eval.Evaluate(evaluator, 0, limits[0], start, new double[] { });
                            theUFSession.Eval.Evaluate(evaluator, 0, limits[1], end, new double[] { });

                            Point3d startPoint = new Point3d(start[0], start[1], start[2]);
                            Point3d endPoint = new Point3d(end[0], end[1], end[2]);

                            if (endpointCounts.ContainsKey(startPoint))
                                endpointCounts[startPoint]++;
                            else
                                endpointCounts[startPoint] = 1;

                            if (endpointCounts.ContainsKey(endPoint))
                                endpointCounts[endPoint]++;
                            else
                                endpointCounts[endPoint] = 1;
                        }
                        finally
                        {
                            if (evaluator != IntPtr.Zero)
                                theUFSession.Eval.Free(evaluator);
                        }
                    }
                }

                List<Point3d> uniqueEndpoints = endpointCounts
                    .Where(kvp => kvp.Value == 1)
                    .Select(kvp => kvp.Key)
                    .ToList();

                return uniqueEndpoints;
            }
            catch (Exception ex)
            {
                LogError(nameof(GetUniqueCurveEndpoints), ex);
                throw;
            }
        }

        public class Point3dComparer : IEqualityComparer<Point3d>
        {
            private readonly double tolerance;

            public Point3dComparer(double tolerance)
            {
                this.tolerance = tolerance;
            }

            public bool Equals(Point3d p1, Point3d p2)
            {
                return Math.Abs(p1.X - p2.X) < tolerance &&
                       Math.Abs(p1.Y - p2.Y) < tolerance &&
                       Math.Abs(p1.Z - p2.Z) < tolerance;
            }

            public int GetHashCode(Point3d point)
            {
                return (point.X.ToString("F6") + point.Y.ToString("F6") + point.Z.ToString("F6")).GetHashCode();
            }
        }

        public static double GetCurveLength(Curve curve)
        {
            try
            {
                if (curve == null)
                    throw new ArgumentNullException(nameof(curve));

                return curve.GetLength();
            }
            catch (Exception ex)
            {
                LogError(nameof(GetCurveLength), ex);
                return 0.0;
            }
        }

        private static double GetTotalCurveLength(Curve[] curves)
        {
            double totalLength = 0.0;
            if (curves == null) return 0.0;

            foreach (Curve curve in curves)
            {
                totalLength += GetCurveLength(curve);
            }

            return totalLength;
        }

        public static double GetOffsetCurveLength(NXObject feature)
        {
            double totalLength = 0.0;

            try
            {
                if (feature is Feature offsetFeature)
                {
                    NXObject[] resultingEntities = offsetFeature.GetEntities();
                    foreach (NXObject entity in resultingEntities)
                    {
                        if (entity is Curve curve)
                        {
                            totalLength += GetCurveLength(curve);
                        }
                    }
                }

                return totalLength;
            }
            catch (Exception ex)
            {
                LogError(nameof(GetOffsetCurveLength), ex);
                return 0.0;
            }
        }

        public static double CalculateDistanceBetweenSketchesAndPanel(List<Sketch> sketches, Body panelFace)
        {
            try
            {
                if (sketches == null || sketches.Count == 0)
                    throw new ArgumentException("Sketches list cannot be null or empty.", nameof(sketches));

                if (panelFace == null)
                    throw new ArgumentNullException(nameof(panelFace));

                Tag panelFaceTag = panelFace.Tag;
                double overallMinDistance = double.MaxValue;

                foreach (Sketch sketch in sketches)
                {
                    NXObject[] sketchCurves = sketch.GetAllGeometry();

                    foreach (NXObject obj in sketchCurves)
                    {
                        if (obj is Curve curve)
                        {
                            Tag curveTag = curve.Tag;
                            double[] guess1 = new double[3];
                            double[] guess2 = new double[3];
                            double minDistance;
                            double[] pointOnCurve = new double[3];
                            double[] pointOnFace = new double[3];

                            theUFSession.Modl.AskMinimumDist(
                                curveTag,
                                panelFaceTag,
                                0,
                                guess1,
                                0,
                                guess2,
                                out minDistance,
                                pointOnCurve,
                                pointOnFace);

                            if (minDistance < overallMinDistance)
                                overallMinDistance = minDistance;
                        }
                    }
                }

                return overallMinDistance;
            }
            catch (Exception ex)
            {
                LogError(nameof(CalculateDistanceBetweenSketchesAndPanel), ex);
                throw;
            }
        }

        public static Point3d ExtractLowestPointFromCurve(Curve curve)
        {
            if (curve == null)
                throw new ArgumentException("Curve cannot be null.", nameof(curve));

            List<Point3d> points = new List<Point3d>();

            try
            {
                LogInfo($"Processing curve with Tag: {curve.Tag}");

                double[] paramRange = new double[2];
                int periodicity;
                theUFSession.Curve.AskParameterization(curve.Tag, paramRange, out periodicity);

                const int numPoints = 150;
                double[] posAndDeriv = new double[9];

                for (int i = 0; i <= numPoints; i++)
                {
                    double parameter = paramRange[0] + i * (paramRange[1] - paramRange[0]) / numPoints;

                    try
                    {
                        theUFSession.Curve.EvaluateCurve(curve.Tag, parameter, 1, posAndDeriv);
                        Point3d point3d = new Point3d(posAndDeriv[0], posAndDeriv[1], posAndDeriv[2]);
                        points.Add(point3d);
                    }
                    catch (Exception ex)
                    {
                        LogError($"{nameof(ExtractLowestPointFromCurve)} EvaluateCurve", ex);
                    }
                }

                Point3d lowestPoint = points.OrderBy(p => p.Z).FirstOrDefault();
                LogInfo($"Lowest Point: X={lowestPoint.X}, Y={lowestPoint.Y}, Z={lowestPoint.Z}");
                return lowestPoint;
            }
            catch (Exception ex)
            {
                LogError(nameof(ExtractLowestPointFromCurve), ex);
                throw;
            }
        }

        #endregion

        #region === Selection & Sketch Utilities ===

        public static Sketch[] SelectSketchesFromUI(string message)
        {
            List<Sketch> selectedSketches = new List<Sketch>();

            try
            {
                UI theUI = UI.GetUI();
                TaggedObject[] selectedObjects;

                NXOpen.Select.FilterMember[] sketchFilter = new NXOpen.Select.FilterMember[1];
                sketchFilter[0] = NXOpen.Select.FilterMember.Sketch;

                Selection.Response response = theUI.SelectionManager.SelectTaggedObjectsWithFilterMembers(
                    message ?? "Select one or more sketches for operation",
                    message ?? "Select sketches",
                    Selection.SelectionScope.WorkPart,
                    Selection.SelectionAction.ClearAndEnableSpecific,
                    sketchFilter,
                    out selectedObjects);

                if (response == Selection.Response.Ok && selectedObjects != null)
                {
                    foreach (TaggedObject obj in selectedObjects)
                    {
                        if (obj is Sketch sketch)
                        {
                            selectedSketches.Add(sketch);
                        }
                    }

                    if (selectedSketches.Count == 0)
                    {
                        ShowMessage("No valid sketches were selected.");
                    }
                }
                else
                {
                    ShowMessage("No sketches were selected.");
                }
            }
            catch (Exception ex)
            {
                ShowMessage("Error during sketch selection: " + ex.Message);
                LogError(nameof(SelectSketchesFromUI), ex);
            }

            return selectedSketches.ToArray();
        }

        public static List<NXObject> ConvertCurveArrayToList(Curve[] curves)
        {
            try
            {
                List<NXObject> nxObjectList = new List<NXObject>();

                if (curves == null)
                    return nxObjectList;

                foreach (Curve curve in curves)
                {
                    if (curve != null)
                        nxObjectList.Add(curve);
                }

                return nxObjectList;
            }
            catch (Exception ex)
            {
                LogError(nameof(ConvertCurveArrayToList), ex);
                throw;
            }
        }

        #endregion

        #region === Offset Curve Utilities ===

        /// <summary>
        /// Single offset distance utility.
        /// </summary>
        public static List<Curve> CreateOffsetCurves(Curve[] curvesToOffset, double offsetDistance)
        {
            List<Curve> offsetCurves = new List<Curve>();
            OffsetCurveBuilder offsetCurveBuilder = null;
            Session session = Session.GetSession();
            Session.UndoMarkId undoMarkId;

            try
            {
                if (curvesToOffset == null || curvesToOffset.Length == 0)
                    throw new ArgumentException("curvesToOffset cannot be null or empty.", nameof(curvesToOffset));

                offsetCurveBuilder = workPart.Features.CreateOffsetCurveBuilder(null);

                offsetCurveBuilder.OffsetDistance.Value = offsetDistance;
                offsetCurveBuilder.CurveFitData.Tolerance = 0.01;
                offsetCurveBuilder.CurveFitData.AngleTolerance = 0.5;
                offsetCurveBuilder.TrimMethod = OffsetCurveBuilder.TrimOption.ExtendTangents;

                SelectionIntentRule[] curveRules = curvesToOffset
                    .Select(curve => workPart.ScRuleFactory.CreateRuleCurveDumb(new Curve[] { curve }))
                    .ToArray();

                offsetCurveBuilder.CurvesToOffset.Clear();
                offsetCurveBuilder.CurvesToOffset.AddToSection(
                    curveRules,
                    curvesToOffset[0],
                    null,
                    null,
                    new Point3d(),
                    Section.Mode.Create,
                    false);

                undoMarkId = session.SetUndoMark(Session.MarkVisibility.Visible, "Offset Curve Commit");

                offsetCurveBuilder.ReverseDirection = false;
                offsetCurveBuilder.RoughOffset = true;
                NXObject offsetFeature = offsetCurveBuilder.CommitFeature();
                double offsetLength = GetOffsetCurveLength(offsetFeature);

                if (offsetLength < curvesToOffset.Sum(curve => GetCurveLength(curve)))
                {
                    session.UndoToMark(undoMarkId, "Undo Offset Commit");

                    offsetCurveBuilder.Destroy();
                    offsetCurveBuilder = workPart.Features.CreateOffsetCurveBuilder(null);
                    offsetCurveBuilder.OffsetDistance.Value = offsetDistance;
                    offsetCurveBuilder.CurveFitData.Tolerance = 0.01;
                    offsetCurveBuilder.CurveFitData.AngleTolerance = 0.5;
                    offsetCurveBuilder.TrimMethod = OffsetCurveBuilder.TrimOption.ExtendTangents;

                    offsetCurveBuilder.ReverseDirection = true;
                    offsetCurveBuilder.CurvesToOffset.Clear();
                    offsetCurveBuilder.CurvesToOffset.AddToSection(
                        curveRules,
                        curvesToOffset[0],
                        null,
                        null,
                        new Point3d(),
                        Section.Mode.Create,
                        false);
                    offsetCurveBuilder.RoughOffset = true;

                    offsetFeature = offsetCurveBuilder.CommitFeature();
                }

                if (offsetFeature is Feature feature)
                {
                    NXObject[] resultingEntities = feature.GetEntities();
                    foreach (NXObject entity in resultingEntities)
                    {
                        if (entity is Curve resultingCurve)
                        {
                            offsetCurves.Add(resultingCurve);
                        }
                    }
                }

                session.DeleteUndoMark(undoMarkId, null);
                LogInfo("CreateOffsetCurves (single) successful.");

                return offsetCurves;
            }
            catch (Exception ex)
            {
                LogError(nameof(CreateOffsetCurves), ex);
                return offsetCurves;
            }
            finally
            {
                offsetCurveBuilder?.Destroy();
            }
        }

        /// <summary>
        /// Creates offsets for multiple distances (e.g., 50 and 10 mm) and returns a mapping.
        /// </summary>
        public static Dictionary<double, List<Curve>> CreateOffsetCurves(Part wp, Curve[] curvesToOffset)
        {
            double[] offsetDistances = { 50.0, 10.0 };
            var resultOffsets = new Dictionary<double, List<Curve>>();
            Session session = Session.GetSession();

            foreach (double offsetDistance in offsetDistances)
            {
                List<Curve> offsetCurves = new List<Curve>();
                OffsetCurveBuilder offsetCurveBuilder = null;
                Session.UndoMarkId undoMarkId;

                try
                {
                    if (curvesToOffset == null || curvesToOffset.Length == 0)
                        throw new ArgumentException("curvesToOffset cannot be null or empty.", nameof(curvesToOffset));

                    offsetCurveBuilder = wp.Features.CreateOffsetCurveBuilder(null);
                    offsetCurveBuilder.OffsetDistance.Value = offsetDistance;
                    offsetCurveBuilder.CurveFitData.Tolerance = 0.01;
                    offsetCurveBuilder.CurveFitData.AngleTolerance = 0.5;
                    offsetCurveBuilder.TrimMethod = OffsetCurveBuilder.TrimOption.ExtendTangents;

                    List<SelectionIntentRule> selectionRules = new List<SelectionIntentRule>();
                    foreach (Curve selectedCurve in curvesToOffset)
                    {
                        SelectionIntentRule curveRule =
                            wp.ScRuleFactory.CreateRuleCurveDumb(new Curve[] { selectedCurve });
                        selectionRules.Add(curveRule);
                    }

                    offsetCurveBuilder.CurvesToOffset.Clear();
                    offsetCurveBuilder.CurvesToOffset.AddToSection(
                        selectionRules.ToArray(),
                        curvesToOffset[0],
                        null,
                        null,
                        new Point3d(),
                        Section.Mode.Create,
                        false);

                    undoMarkId = session.SetUndoMark(Session.MarkVisibility.Visible, $"Offset {offsetDistance} Commit");

                    NXObject offsetFeature = null;

                    if (offsetDistance == 50.0)
                    {
                        offsetCurveBuilder.ReverseDirection = true;
                        offsetFeature = offsetCurveBuilder.CommitFeature();
                        double offsetLength = GetOffsetCurveLength(offsetFeature);

                        if (offsetLength <= GetTotalCurveLength(curvesToOffset))
                        {
                            session.UndoToMark(undoMarkId, $"Undo Offset {offsetDistance} Commit");

                            offsetCurveBuilder.Destroy();
                            offsetCurveBuilder = wp.Features.CreateOffsetCurveBuilder(null);
                            offsetCurveBuilder.OffsetDistance.Value = offsetDistance;
                            offsetCurveBuilder.CurveFitData.Tolerance = 0.01;
                            offsetCurveBuilder.CurveFitData.AngleTolerance = 0.5;
                            offsetCurveBuilder.TrimMethod = OffsetCurveBuilder.TrimOption.ExtendTangents;

                            offsetCurveBuilder.ReverseDirection = false;
                            offsetCurveBuilder.CurvesToOffset.Clear();
                            offsetCurveBuilder.CurvesToOffset.AddToSection(
                                selectionRules.ToArray(),
                                curvesToOffset[0],
                                null,
                                null,
                                new Point3d(),
                                Section.Mode.Create,
                                false);
                            offsetFeature = offsetCurveBuilder.CommitFeature();
                        }
                    }

                    if (offsetDistance == 10.0)
                    {
                        offsetCurveBuilder.ReverseDirection = false;
                        offsetFeature = offsetCurveBuilder.CommitFeature();
                        double offsetLength = GetOffsetCurveLength(offsetFeature);

                        if (offsetLength >= GetTotalCurveLength(curvesToOffset))
                        {
                            session.UndoToMark(undoMarkId, $"Undo Offset {offsetDistance} Commit");

                            offsetCurveBuilder.Destroy();
                            offsetCurveBuilder = wp.Features.CreateOffsetCurveBuilder(null);
                            offsetCurveBuilder.OffsetDistance.Value = offsetDistance;
                            offsetCurveBuilder.CurveFitData.Tolerance = 0.01;
                            offsetCurveBuilder.CurveFitData.AngleTolerance = 0.5;
                            offsetCurveBuilder.TrimMethod = OffsetCurveBuilder.TrimOption.ExtendTangents;

                            offsetCurveBuilder.ReverseDirection = true;
                            offsetCurveBuilder.CurvesToOffset.Clear();
                            offsetCurveBuilder.CurvesToOffset.AddToSection(
                                selectionRules.ToArray(),
                                curvesToOffset[0],
                                null,
                                null,
                                new Point3d(),
                                Section.Mode.Create,
                                false);
                            offsetFeature = offsetCurveBuilder.CommitFeature();
                        }
                    }

                    if (offsetFeature is Feature feature)
                    {
                        NXObject[] resultingEntities = feature.GetEntities();
                        foreach (NXObject entity in resultingEntities)
                        {
                            if (entity is Curve resultingCurve)
                            {
                                offsetCurves.Add(resultingCurve);
                            }
                        }
                    }

                    session.DeleteUndoMark(undoMarkId, null);
                }
                catch (Exception ex)
                {
                    LogError($"{nameof(CreateOffsetCurves)}[{offsetDistance}]", ex);
                }
                finally
                {
                    offsetCurveBuilder?.Destroy();
                }

                resultOffsets[offsetDistance] = offsetCurves;
            }

            LogInfo("CreateOffsetCurves (multi distance) completed.");
            return resultOffsets;
        }

        #endregion

        #region === Datum & Projection Utilities ===

        public static DatumPlane CreateDatum(DatumPlane referenceDatumPlane, Point3d targetPoint)
        {
            Session.UndoMarkId markId = theSession.SetUndoMark(Session.MarkVisibility.Visible, "Create Datum Plane");
            DatumPlaneBuilder datumPlaneBuilder = null;
            Point point = null;

            try
            {
                if (referenceDatumPlane == null)
                    throw new ArgumentNullException(nameof(referenceDatumPlane));

                datumPlaneBuilder = workPart.Features.CreateDatumPlaneBuilder(null);
                Plane plane = datumPlaneBuilder.GetPlane();

                plane.SetMethod(PlaneTypes.MethodType.ParallelPoint);

                point = workPart.Points.CreatePoint(targetPoint);
                NXObject[] geometry = { referenceDatumPlane, point };
                plane.SetGeometry(geometry);

                plane.SetOffsetExpression("-40");
                plane.SetUpdateOption(SmartObject.UpdateOption.WithinModeling);
                plane.Evaluate();

                Feature feature = datumPlaneBuilder.CommitFeature();
                DatumPlaneFeature datumPlaneFeature = feature as DatumPlaneFeature;

                if (datumPlaneFeature == null)
                    throw new InvalidCastException("The created feature is not a DatumPlaneFeature.");

                DatumPlane createdDatumPlane = datumPlaneFeature.DatumPlane;
                LogInfo("CreateDatum successful.");

                return createdDatumPlane;
            }
            catch (Exception ex)
            {
                LogError(nameof(CreateDatum), ex);
                return null;
            }
            finally
            {
                if (point != null)
                {
                    try { workPart.Points.DeletePoint(point); } catch { }
                }

                datumPlaneBuilder?.Destroy();
                theSession.DeleteUndoMark(markId, null);
            }
        }

        public static Curve[] ProjectCurvesAndEdges(Part wp, List<NXObject> curvesAndEdges, DatumPlane planeToProjectTo)
        {
            Session.UndoMarkId markId1 = theSession.SetUndoMark(Session.MarkVisibility.Visible, "Project Curves");
            ProjectCurveBuilder projectCurveBuilder = null;

            try
            {
                if (wp == null)
                    throw new ArgumentNullException(nameof(wp));

                if (curvesAndEdges == null || curvesAndEdges.Count == 0)
                    throw new ArgumentException("curvesAndEdges cannot be null or empty.", nameof(curvesAndEdges));

                if (planeToProjectTo == null)
                    throw new ArgumentNullException(nameof(planeToProjectTo));

                projectCurveBuilder = wp.Features.CreateProjectCurveBuilder(null);

                projectCurveBuilder.CurveFitData.Tolerance = 0.01;
                projectCurveBuilder.CurveFitData.AngleTolerance = 0.5;
                projectCurveBuilder.ProjectionDirectionMethod = ProjectCurveBuilder.DirectionType.AlongFaceNormal;
                projectCurveBuilder.ProjectionOption = ProjectCurveBuilder.ProjectionOptionType.EqualArcLength;

                projectCurveBuilder.SectionToProject.DistanceTolerance = 0.01;
                projectCurveBuilder.SectionToProject.ChainingTolerance = 0.0095;
                projectCurveBuilder.SectionToProject.AngleTolerance = 0.5;
                projectCurveBuilder.SectionToProject.AllowSelfIntersection(true);
                projectCurveBuilder.SectionToProject.AllowDegenerateCurves(false);
                projectCurveBuilder.SectionToProject.SetAllowedEntityTypes(Section.AllowTypes.CurvesAndPoints);

                List<SelectionIntentRule> selectionRules = new List<SelectionIntentRule>();
                SelectionIntentRuleOptions ruleOptions = wp.ScRuleFactory.CreateRuleOptions();
                ruleOptions.SetSelectedFromInactive(false);

                foreach (NXObject obj in curvesAndEdges)
                {
                    if (obj is Curve curve)
                    {
                        SelectionIntentRule tangentRule = wp.ScRuleFactory.CreateRuleCurveTangent(
                            curve,
                            null,
                            true,
                            0.5,
                            0.01);
                        selectionRules.Add(tangentRule);
                    }
                    else if (obj is Edge edge)
                    {
                        SelectionIntentRule rule = wp.ScRuleFactory.CreateRuleEdgeTangent(
                            edge,
                            null,
                            true,
                            0.5,
                            false,
                            false,
                            ruleOptions);
                        selectionRules.Add(rule);
                    }
                }

                projectCurveBuilder.SectionToProject.AddToSection(
                    selectionRules.ToArray(),
                    curvesAndEdges[0],
                    null,
                    null,
                    new Point3d(),
                    Section.Mode.Create,
                    false);

                ScCollector scCollector1 = wp.ScCollectors.CreateCollector();
                SelectionIntentRuleOptions selectionIntentRuleOptions1 = wp.ScRuleFactory.CreateRuleOptions();
                selectionIntentRuleOptions1.SetSelectedFromInactive(false);

                DatumPlane[] faces1 = { planeToProjectTo };
                FaceDumbRule faceDumbRule1 =
                    wp.ScRuleFactory.CreateRuleFaceDatum(faces1, selectionIntentRuleOptions1);

                selectionIntentRuleOptions1.Dispose();
                SelectionIntentRule[] rules2 = { faceDumbRule1 };
                scCollector1.ReplaceRules(rules2, false);

                projectCurveBuilder.FaceToProjectTo.Add(scCollector1);

                NXObject[] projectedObjects = projectCurveBuilder.CommitFeature().GetEntities();
                List<Curve> projectedCurves = new List<Curve>();

                foreach (NXObject projectedObject in projectedObjects)
                {
                    if (projectedObject is Curve projectedCurve)
                    {
                        projectedCurves.Add(projectedCurve);
                    }
                }

                LogInfo("ProjectCurvesAndEdges successful.");
                return projectedCurves.ToArray();
            }
            catch (Exception ex)
            {
                LogError(nameof(ProjectCurvesAndEdges), ex);
                return Array.Empty<Curve>();
            }
            finally
            {
                if (projectCurveBuilder != null)
                {
                    try
                    {
                        projectCurveBuilder.SectionToProject.CleanMappingData();
                    }
                    catch
                    {
                    }

                    projectCurveBuilder.Destroy();
                }

                theSession.DeleteUndoMark(markId1, null);
            }
        }

        public static Curve ConvertNXObjectToCurve(NXObject nxObject)
        {
            try
            {
                if (nxObject is Curve curve)
                {
                    return curve;
                }

                throw new InvalidCastException("The provided NXObject is not a Curve.");
            }
            catch (Exception ex)
            {
                LogError(nameof(ConvertNXObjectToCurve), ex);
                throw;
            }
        }

        public static Curve[] CreateLineBetweenPoints(Point3d startPoint, Point3d endPoint)
        {
            try
            {
                Line createdLine = workPart.Curves.CreateLine(startPoint, endPoint);
                return new Curve[] { createdLine };
            }
            catch (Exception ex)
            {
                LogError(nameof(CreateLineBetweenPoints), ex);
                throw new Exception("Error while creating line: " + ex.Message, ex);
            }
        }

        #endregion

        #region === General Helpers ===

        public static void ShowMessage(string message)
        {
            try
            {
                UI theUI = UI.GetUI();
                theUI.NXMessageBox.Show("Message", NXMessageBox.DialogType.Information, message);
            }
            catch (Exception ex)
            {
                LogError(nameof(ShowMessage), ex);
            }
        }

        public static int GetUnloadOption(string dummy)
        {
            return (int)Session.LibraryUnloadOption.Immediately;
        }

        #endregion
    }
}
