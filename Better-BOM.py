#Author-CyberReefGuru
#Description-Updated version of CSV BOM that creates better outcomes.  Thanks to Peter Boeker for doing the heavy lifting.  This would not be possible without his work!

import adsk.core, adsk.fusion, adsk.cam, traceback, json, re

# Global list to keep all event handlers in scope.
# This is only needed with Python.
handlers = []
app = adsk.core.Application.get()
ui = app.userInterface
cmdId = "BillOfMaterialAddInMenuEntry"
cmdName = "Better-BOM"
dialogTitle = "Create a Bill of Materials"
cmdDesc = "Creates a bill of material from components within your design."
cmdRes = ".//resources"
cmdToolbarLocation = "SolidCreatePanel"

KEY_PREF = "lastUsedOptions"
KEY_PREF_ONLY_SELECTED = "onlySelComp"
KEY_PREF_INC_DIMENSIONS = "incBoundDims"
KEY_PREF_SORT_DIM = "sortDims"
KEY_PREF_IGNORE_PREFIXED_COMP = "ignoreUnderscorePrefComp"
KEY_PREF_STRIP_UNDERSCORE = "underscorePrefixStrip"
KEY_PREF_IGNORE_NO_BODIES = "ignoreCompWoBodies"
KEY_PREF_IGNORE_LINKED = "ignoreLinkedComp"
KEY_PREF_IGNORE_INVISIBLE = "ignoreVisibleState"
KEY_PREF_INC_PART_NUMBER = "incPartNumber"
KEY_PREF_INC_VOLUME = "incVol"
KEY_PREF_INC_AREA = "incArea"
KEY_PREF_INC_MASS = "incMass"
KEY_PREF_INC_DENSITY = "incDensity"
KEY_PREF_INC_MATERIAL = "incMaterial"
KEY_PREF_INC_PARENT = "incParent"
KEY_PREF_INC_DESCRIPTION = "incDesc"
KEY_PREF_GEN_CUT_LIST = "generateCutList"
KEY_PREF_USE_COMMA = "useComma"

# Event handler for the commandCreated event. Called when this addin is initially created.
class BOMCommandCreatedEventHandler(adsk.core.CommandCreatedEventHandler):
    global cmdId
    global ui

    # init method required for event handlers
    def __init__(self):
        super().__init__()

    # notify method called when command created event is fired
    def notify(self, args):

        # Initialized preference variables; previously set values will be used if they exist
        product = app.activeProduct
        design = adsk.fusion.Design.cast(product)
        lastPrefs = design.attributes.itemByName(cmdId, "lastUsedOptions")
        _onlySelectedComps = False
        _includeBoundingboxDims = True
        _sortDims = False
        _ignoreUnderscorePrefixedComps = False
        _underscorePrefixStrip = False
        _ignoreCompsWithoutBodies = True
        _ignoreLinkedComps = False
        _ignoreVisibleState = True
        _includePartNumber = True
        _includeVolume = False
        _includeArea = False
        _includeMass = False
        _includeDensity = False
        _includeMaterial = False
        _includeParent = False
        _generateCutList = False
        _includeDesc = False
        _useComma = False

        # load preferences from file system if it exists; if pref doesn't exist, set default expressed above.
        if lastPrefs:
            try:
                lastPrefs = json.loads(lastPrefs.value)
                _onlySelectedComps = lastPrefs.get(KEY_PREF_ONLY_SELECTED, _onlySelectedComps)
                _includeBoundingboxDims = lastPrefs.get(KEY_PREF_INC_DIMENSIONS, _includeBoundingboxDims)
                _sortDims = lastPrefs.get(KEY_PREF_SORT_DIM, _sortDims)
                _ignoreUnderscorePrefixedComps = lastPrefs.get(KEY_PREF_IGNORE_PREFIXED_COMP, _ignoreUnderscorePrefixedComps)
                _underscorePrefixStrip = lastPrefs.get(KEY_PREF_STRIP_UNDERSCORE, _underscorePrefixStrip)
                _ignoreCompsWithoutBodies = lastPrefs.get(KEY_PREF_IGNORE_NO_BODIES, _ignoreCompsWithoutBodies)
                _ignoreLinkedComps = lastPrefs.get(KEY_PREF_IGNORE_LINKED, _ignoreLinkedComps)
                _ignoreVisibleState = lastPrefs.get(KEY_PREF_IGNORE_INVISIBLE, _ignoreVisibleState)
                _includePartNumber = lastPrefs.get(KEY_PREF_INC_PART_NUMBER, _includePartNumber)
                _includeVolume = lastPrefs.get(KEY_PREF_INC_VOLUME, _includeVolume)
                _includeArea = lastPrefs.get(KEY_PREF_INC_AREA, _includeArea)
                _includeMass = lastPrefs.get(KEY_PREF_INC_MASS, _includeMass)
                _includeDensity = lastPrefs.get(KEY_PREF_INC_DENSITY, _includeDensity)
                _includeMaterial = lastPrefs.get(KEY_PREF_INC_MATERIAL, _includeMaterial)
                _includeParent = lastPrefs.get(KEY_PREF_INC_PARENT, _includeParent)
                _generateCutList = lastPrefs.get(KEY_PREF_GEN_CUT_LIST, _generateCutList)
                _includeDesc = lastPrefs.get(KEY_PREF_INC_DESCRIPTION, _includeDesc)
                _useComma = lastPrefs.get(KEY_PREF_USE_COMMA, _useComma)
            except:
                ui.messageBox('Failed to load preferences:\n{}'.format(traceback.format_exc()))
                return

        # get the arguments for this event
        eventArgs = adsk.core.CommandCreatedEventArgs.cast(args)
        
        # get the command assoicated with this event
        cmd = eventArgs.command
        
        # get the inputs associated with this command
        inputs = cmd.commandInputs
        ipSelectComps = inputs.addBoolValueInput(cmdId + KEY_PREF_ONLY_SELECTED, "Include only selected components", True, "", _onlySelectedComps)
        ipSelectComps.tooltip = "Only selected components will be used."

        ipBoundingBox = inputs.addBoolValueInput(cmdId + KEY_PREF_INC_DIMENSIONS, "Include dimensions", True, "", _includeBoundingboxDims)
        ipBoundingBox.tooltip = "Includes dimensions of bodies."

        ipCompDesc = inputs.addBoolValueInput(cmdId + KEY_PREF_INC_DESCRIPTION, "Include description", True, "", _includeDesc)
        ipCompDesc.tooltip = "Includes the component description. You can add a description<br/>by right clicking a component and open the Properties panel."
      
        # Group additional properties into a submenu
        grpPhysics = inputs.addGroupCommandInput(cmdId + "_grpPhysics", "Additional Properties")
        if _includeVolume or _includeArea or _includeMass or _includeDensity or _includeMaterial:
        	grpPhysics.isExpanded = True
        else:
        	grpPhysics.isExpanded = False
        grpPhysicsChildren = grpPhysics.children

        ipVolume = grpPhysicsChildren.addBoolValueInput(cmdId + KEY_PREF_INC_VOLUME, "Include volume", True, "", _includeVolume)
        ipVolume.tooltip = "Adds the calculated volume of all bodies related to the parent component"

        ipIncludeArea = grpPhysicsChildren.addBoolValueInput(cmdId + KEY_PREF_INC_AREA, "Include area", True, "", _includeArea)
        ipIncludeArea.tooltip = "Include component area in cm^2"

        ipIncludeMass = grpPhysicsChildren.addBoolValueInput(cmdId + KEY_PREF_INC_MASS, "Include mass", True, "", _includeMass)
        ipIncludeMass.tooltip = "Include component mass in kg"

        ipIncludeDensity = grpPhysicsChildren.addBoolValueInput(cmdId + KEY_PREF_INC_DENSITY, "Include density", True, "", _includeDensity)
        ipIncludeDensity.tooltip = "Include component density in kg/cm^3"

        ipIncludeMaterial = grpPhysicsChildren.addBoolValueInput(cmdId + KEY_PREF_INC_MATERIAL, "Include material", True, "", _includeMaterial)
        ipIncludeMaterial.tooltip = "Include component physical material"

        # Group advanced options into a submenu
        grpMisc = inputs.addGroupCommandInput(cmdId + "_grpAdvanced", "Advanced")
        grpMisc.isExpanded = False
        grpMisc.isVisible = True
        grpMiscChildren = grpMisc.children

        ipGenerateCutList = grpMiscChildren.addBoolValueInput(cmdId + KEY_PREF_GEN_CUT_LIST, "Generate Cut List", True, "", _generateCutList)
        ipGenerateCutList.tooltip = "Generates cut list for body dimensions."

        ipUseComma = grpMiscChildren.addBoolValueInput(cmdId + KEY_PREF_USE_COMMA, "Use Comma Delimiter", True, "", _useComma)
        ipUseComma.tooltip = "Uses comma instead of point for number decimal delimiter."

        ipParent = grpMiscChildren.addBoolValueInput(cmdId + KEY_PREF_INC_PARENT, "Include Parent Component", True, "", _includeParent)
        ipParent.tooltip = "Include the parent component in BOM output."

        ipPart = grpMiscChildren.addBoolValueInput(cmdId + KEY_PREF_INC_PART_NUMBER, "Include Part Number", True, "", _includePartNumber)
        ipPart.tooltip = "Include the component part number in BOM output."

        ipWoBodies = grpMiscChildren.addBoolValueInput(cmdId + KEY_PREF_IGNORE_NO_BODIES, "Ignore Components without Bodies", True, "", _ignoreCompsWithoutBodies)
        ipWoBodies.tooltip = "Exclude component if it has no bodies."

        ipLinkedComps = grpMiscChildren.addBoolValueInput(cmdId + KEY_PREF_IGNORE_LINKED, "Ignore External Components", True, "", _ignoreLinkedComps)
        ipLinkedComps.tooltip = "Exclude external components."

        ipVisibleState = grpMiscChildren.addBoolValueInput(cmdId + KEY_PREF_IGNORE_INVISIBLE, "Ignore Invisible Components", True, "", _ignoreVisibleState)
        ipVisibleState.tooltip = "Ignores/excludes components that are not visible in the design."

        ipsortDims = grpMiscChildren.addBoolValueInput(cmdId + KEY_PREF_SORT_DIM, "Sort dimensions", True, "", _sortDims)
        ipsortDims.tooltip = "Sorts the dimensions of bodies - smallest dimension is the height (thickness), next larger is the width, largest is the length."
        ipsortDims.isVisible = _includeBoundingboxDims

        ipUnderscorePrefix = grpMiscChildren.addBoolValueInput(cmdId + KEY_PREF_IGNORE_PREFIXED_COMP, 'Exclude "_"', True, "", _ignoreUnderscorePrefixedComps)
        ipUnderscorePrefix.tooltip = 'Exclude all components there name starts with "_"'

        ipUnderscorePrefixStrip = grpMiscChildren.addBoolValueInput(cmdId + KEY_PREF_STRIP_UNDERSCORE, 'Strip "_"', True, "", _underscorePrefixStrip)
        ipUnderscorePrefixStrip.tooltip = 'If checked, "_" is stripped from components name'

        # Connect to the execute event.
        onExecute = BOMCommandExecuteHandler()
        cmd.execute.add(onExecute)
        handlers.append(onExecute)

        # onInputChanged = BOMCommandInputChangedHandler()
        # cmd.inputChanged.add(onInputChanged)
        # handlers.append(onInputChanged)


# Event handler for the execute event.
class BOMCommandExecuteHandler(adsk.core.CommandEventHandler):
    global cmdId
    def __init__(self):
        super().__init__()

    def replacePointDelimterOnPref(self, pref, value):
        if (pref):
            return str(value).replace(".", ",")
        return str(value)

    # Collects all the BOM information and builds the BOM output file
    def collectData(self, design, bom, prefs):
        csvStr = ''
        defaultUnit = design.fusionUnitsManager.defaultLengthUnits
        csvHeader = ["Part Name"]

        if prefs[KEY_PREF_INC_PARENT]:
            csvHeader.append("Parent")

        if prefs[KEY_PREF_INC_PART_NUMBER]:
            csvHeader.append("Part Number")

        csvHeader.append("Quantity")

        if prefs[KEY_PREF_INC_DESCRIPTION]:
            csvHeader.append("Description")
        
        if prefs[KEY_PREF_INC_DIMENSIONS]:
        		csvHeader.append("Width " + defaultUnit)
        		csvHeader.append("Length " + defaultUnit)
        		csvHeader.append("Height " + defaultUnit)

        #
        # TODO: units here need to be adjusted based on default units
        # 
        if prefs[KEY_PREF_INC_VOLUME]:
            csvHeader.append("Volume cm^3")
        if prefs[KEY_PREF_INC_AREA]:
            csvHeader.append("Area cm^2")
        if prefs[KEY_PREF_INC_MASS]:
            csvHeader.append("Mass kg")
        if prefs[KEY_PREF_INC_DENSITY]:
            csvHeader.append("Density kg/cm^2")

        if prefs[KEY_PREF_INC_MATERIAL]:
        	csvHeader.append("Material")
        for k in csvHeader:
            csvStr += '"' + k + '",'
        csvStr += '\n'

        # for each unique component found, write to string
        for item in bom:
            # Name, Parent, Part #, Qtny, Desc, W, L, H, Vol, Area, Mass, Den, Mat
            dims = ''
            name = self.filterFusionCompNameInserts(item["name"])
            if prefs[KEY_PREF_IGNORE_PREFIXED_COMP] is False and prefs[KEY_PREF_STRIP_UNDERSCORE] is True and name[0] == '_':
                name = name[1:]

            # write name
            csvStr += '"' + name + '",'
            
            # write parent if option selected
            if prefs[KEY_PREF_INC_PARENT]:
                csvStr += '"' + item["parent"] + '",'

            # write part number if option selected
            if prefs[KEY_PREF_INC_PART_NUMBER]:
                csvStr += '"' + item["partnumber"] + '",'

            # Write Quantity
            csvStr += '"' + self.replacePointDelimterOnPref(prefs[KEY_PREF_USE_COMMA], item["instances"]) + '",'
            
            # write description if option selected
            if prefs[KEY_PREF_INC_DESCRIPTION]:
                csvStr += '"' + item["desc"] + '",'

            # write dimension information if option selected
            if prefs[KEY_PREF_INC_DIMENSIONS]:
            	dim = 0
            	for k in item["boundingBox"]:
            		dim += item["boundingBox"][k]
            	if dim > 0:
            		dimX = float(design.fusionUnitsManager.formatInternalValue(item["boundingBox"]["x"], defaultUnit, False))
            		dimY = float(design.fusionUnitsManager.formatInternalValue(item["boundingBox"]["y"], defaultUnit, False))
            		dimZ = float(design.fusionUnitsManager.formatInternalValue(item["boundingBox"]["z"], defaultUnit, False))
            		if prefs[KEY_PREF_SORT_DIM]:
            			dimSorted = sorted([dimX, dimY, dimZ])
            			bbZ = "{0:.3f}".format(dimSorted[0])
            			bbX = "{0:.3f}".format(dimSorted[1])
            			bbY = "{0:.3f}".format(dimSorted[2])
            		else:
            			bbX = "{0:.3f}".format(dimX)
            			bbY = "{0:.3f}".format(dimY)
            			bbZ = "{0:.3f}".format(dimZ)
                        # write dimensions
            			csvStr += '"' + self.replacePointDelimterOnPref(prefs[KEY_PREF_USE_COMMA], bbX) + '",'
            			csvStr += '"' + self.replacePointDelimterOnPref(prefs[KEY_PREF_USE_COMMA], bbY) + '",'
            			csvStr += '"' + self.replacePointDelimterOnPref(prefs[KEY_PREF_USE_COMMA], bbZ) + '",'
            	else:
                    csvStr += "0" + ','
                    csvStr += "0" + ','
                    csvStr += "0" + ','

            if prefs[KEY_PREF_INC_VOLUME]:
            	csvStr += '"' + self.replacePointDelimterOnPref(prefs[KEY_PREF_USE_COMMA], item["volume"]) + '",'
            if prefs[KEY_PREF_INC_AREA]:
            	csvStr += '"' + self.replacePointDelimterOnPref(prefs[KEY_PREF_USE_COMMA], "{0:.2f}".format(item["area"])) + '",'
            if prefs[KEY_PREF_INC_MASS]:
            	csvStr += '"' + self.replacePointDelimterOnPref(prefs[KEY_PREF_USE_COMMA], "{0:.5f}".format(item["mass"])) + '",'
            if prefs[KEY_PREF_INC_DENSITY]:
            	csvStr += '"' + self.replacePointDelimterOnPref(prefs[KEY_PREF_USE_COMMA], "{0:.5f}".format(item["density"])) + '",'
            if prefs[KEY_PREF_INC_MATERIAL]:
            	csvStr += '"' + item["material"] + '",'

            # write newline
            csvStr += '\n'
        return csvStr

    def collectCutList(self, design, bom, prefs):
        defaultUnit = design.fusionUnitsManager.defaultLengthUnits

        # Init CutList Header
        cutListStr = 'V2\n'
        if prefs[KEY_PREF_USE_COMMA]:
            cutListStr += 'FormatSettings.decimalseparator,\n'
        else:
            cutListStr += 'FormatSettings.decimalseparator.\n'
        cutListStr += '\n'
        cutListStr += 'Required\n'

        #add parts:
        for item in bom:
            name = self.filterFusionCompNameInserts(item["name"])
            if prefs[KEY_PREF_IGNORE_PREFIXED_COMP] is False and prefs[KEY_PREF_STRIP_UNDERSCORE] is True and name[0] == '_':
                name = name[1:]
            # dimensions:
            dim = 0
            for k in item["boundingBox"]:
                dim += item["boundingBox"][k]
            if dim > 0:
                dimX = float(design.fusionUnitsManager.formatInternalValue(item["boundingBox"]["x"], defaultUnit, False))
                dimY = float(design.fusionUnitsManager.formatInternalValue(item["boundingBox"]["y"], defaultUnit, False))
                dimZ = float(design.fusionUnitsManager.formatInternalValue(item["boundingBox"]["z"], defaultUnit, False))

                if prefs["sortDims"]:
                    dims = sorted([dimX, dimY, dimZ])
                else:
                    dims = [dimZ, dimX, dimY]

                partStr = ' '  # leading space
                partStr += self.replacePointDelimterOnPref(prefs[KEY_PREF_USE_COMMA], "{0:.3f}".format(dims[1])).ljust(9)  # width
                partStr += self.replacePointDelimterOnPref(prefs[KEY_PREF_USE_COMMA], "{0:.3f}".format(dims[2])).ljust(7)  # length

                partStr += name
                partStr += ' (thickness: ' + self.replacePointDelimterOnPref(prefs[KEY_PREF_USE_COMMA], "{0:.3f}".format(dims[0])) + defaultUnit + ')'
                partStr += '\n'

            else:
                partStr = ' 0        0      ' + name + '\n'

            # add all instances of the component to the CutList:
            quantity = int(item["instances"])
            for i in range(0, quantity):
                cutListStr += partStr
                

        # empty entry for available materials (sheets):
        cutListStr += '\n' + "Available" + '\n'

        return cutListStr

    def getPrefsObject(self, inputs):
        obj = {
            KEY_PREF_ONLY_SELECTED: inputs.itemById(cmdId + KEY_PREF_ONLY_SELECTED).value,
            KEY_PREF_INC_DIMENSIONS: inputs.itemById(cmdId + KEY_PREF_INC_DIMENSIONS).value,
            KEY_PREF_SORT_DIM: inputs.itemById(cmdId + KEY_PREF_SORT_DIM).value,
            KEY_PREF_IGNORE_PREFIXED_COMP: inputs.itemById(cmdId + KEY_PREF_IGNORE_PREFIXED_COMP).value,
            KEY_PREF_STRIP_UNDERSCORE: inputs.itemById(cmdId + KEY_PREF_STRIP_UNDERSCORE).value,
            KEY_PREF_IGNORE_NO_BODIES: inputs.itemById(cmdId + KEY_PREF_IGNORE_NO_BODIES).value,
            KEY_PREF_IGNORE_LINKED: inputs.itemById(cmdId + KEY_PREF_IGNORE_LINKED).value,
            KEY_PREF_IGNORE_INVISIBLE: inputs.itemById(cmdId + KEY_PREF_IGNORE_INVISIBLE).value,
            KEY_PREF_INC_VOLUME: inputs.itemById(cmdId + KEY_PREF_INC_VOLUME).value,
            KEY_PREF_INC_AREA: inputs.itemById(cmdId + KEY_PREF_INC_AREA).value,
            KEY_PREF_INC_MASS: inputs.itemById(cmdId + KEY_PREF_INC_MASS).value,
            KEY_PREF_INC_DENSITY: inputs.itemById(cmdId + KEY_PREF_INC_DENSITY).value,
            KEY_PREF_INC_MATERIAL: inputs.itemById(cmdId + KEY_PREF_INC_MATERIAL).value,
            KEY_PREF_INC_DESCRIPTION: inputs.itemById(cmdId + KEY_PREF_INC_DESCRIPTION).value,
            KEY_PREF_INC_PARENT: inputs.itemById(cmdId + KEY_PREF_INC_PARENT).value,
            KEY_PREF_INC_PART_NUMBER: inputs.itemById(cmdId + KEY_PREF_INC_PART_NUMBER).value,
            KEY_PREF_GEN_CUT_LIST: inputs.itemById(cmdId + KEY_PREF_GEN_CUT_LIST).value,
            KEY_PREF_USE_COMMA: inputs.itemById(cmdId + KEY_PREF_USE_COMMA).value
        }
        return obj

    def getBodiesVolume(self, bodies):
        volume = 0
        for bodyK in bodies:
            if bodyK.isSolid:
                volume += bodyK.volume
        return volume

    # Calculates a tight bounding box around the input body.  An optional
    # tolerance argument is available.  This specificies the tolerance in
    # centimeters.  If not provided the best existing display mesh is used.
    def calculateTightBoundingBox(self, body, tolerance=0):
        try:
            # If the tolerance is zero, use the best display mesh available.
            if tolerance <= 0:
                # Get the best display mesh available.
                triMesh = body.meshManager.displayMeshes.bestMesh
            else:
                # Calculate a new mesh based on the input tolerance.
                meshMgr = adsk.fusion.MeshManager.cast(body.meshManager)
                meshCalc = meshMgr.createMeshCalculator()
                meshCalc.surfaceTolerance = tolerance
                triMesh = meshCalc.calculate()

            # Calculate the range of the mesh.
            smallPnt = adsk.core.Point3D.cast(triMesh.nodeCoordinates[0])
            largePnt = adsk.core.Point3D.cast(triMesh.nodeCoordinates[0])
            vertex = adsk.core.Point3D.cast(None)
            for vertex in triMesh.nodeCoordinates:
                if vertex.x < smallPnt.x:
                    smallPnt.x = vertex.x

                if vertex.y < smallPnt.y:
                    smallPnt.y = vertex.y

                if vertex.z < smallPnt.z:
                    smallPnt.z = vertex.z

                if vertex.x > largePnt.x:
                    largePnt.x = vertex.x

                if vertex.y > largePnt.y:
                    largePnt.y = vertex.y

                if vertex.z > largePnt.z:
                    largePnt.z = vertex.z

            # Create and return a BoundingBox3D as the result.
            return(adsk.core.BoundingBox3D.create(smallPnt, largePnt))
        except:
            # An error occurred so return None.
            return None

    def getBodiesBoundingBox(self, bodies):
        minPointX = maxPointX = minPointY = maxPointY = minPointZ = maxPointZ = 0
        # Examining the maximum min point distance and the maximum max point distance.
        for body in bodies:
            if body.isSolid:
                bb = self.calculateTightBoundingBox(body, 0)
                if not bb:
                    return None
                if not minPointX or bb.minPoint.x < minPointX:
                    minPointX = bb.minPoint.x
                if not maxPointX or bb.maxPoint.x > maxPointX:
                    maxPointX = bb.maxPoint.x
                if not minPointY or bb.minPoint.y < minPointY:
                    minPointY = bb.minPoint.y
                if not maxPointY or bb.maxPoint.y > maxPointY:
                    maxPointY = bb.maxPoint.y
                if not minPointZ or bb.minPoint.z < minPointZ:
                    minPointZ = bb.minPoint.z
                if not maxPointZ or bb.maxPoint.z > maxPointZ:
                    maxPointZ = bb.maxPoint.z
        return {
            "x": maxPointX - minPointX,
            "y": maxPointY - minPointY,
            "z": maxPointZ - minPointZ
        }

    def getPhysicsArea(self, bodies):
        area = 0
        for body in bodies:
            if body.isSolid:
                if body.physicalProperties:
                    area += body.physicalProperties.area
        return area

    def getPhysicalMass(self, bodies):
        mass = 0
        for body in bodies:
            if body.isSolid:
                if body.physicalProperties:
                    mass += body.physicalProperties.mass
        return mass

    def getPhysicalDensity(self, bodies):
        density = 0
        if bodies.count > 0:
            body = bodies.item(0)
            if body.isSolid:
                if body.physicalProperties:
                    density = body.physicalProperties.density
            return density

    def getPhysicalMaterial(self, bodies):
        matList = []
        for body in bodies:
            if body.isSolid and body.material:
                mat = body.material.name
                if mat not in matList:
                    matList.append(mat)
        return ', '.join(matList)

    def filterFusionCompNameInserts(self, name):
        name = re.sub("\([0-9]+\)$", '', name)
        name = name.strip()
        name = re.sub("v[0-9]+$", '', name)
        return name.strip()


    # Called when plugin is notified of an action
    # This essentially initiatives the plugin's action
    def notify(self, args):
        global app
        global ui
        global dialogTitle
        global cmdId

        product = app.activeProduct
        design = adsk.fusion.Design.cast(product)
        eventArgs = adsk.core.CommandEventArgs.cast(args)
        inputs = eventArgs.command.commandInputs

        _noBody = 0
        _notVisible = 0
        _linked = 0
        _skipped = 0
        _added = 0
        _total = 0

        if not design:
            ui.messageBox('No active design', dialogTitle)
            return

        try:
            prefs = self.getPrefsObject(inputs)

            # Get all occurrences in the root component of the active design
            root = design.rootComponent
            occs = []
            if prefs["onlySelComp"]:
                print("Adding only selected components.")
                if ui.activeSelections.count > 0:
                    selections = ui.activeSelections
                    print("Selected: ", selections.count)
                    for selection in selections:
                        # Loop through everything that is selected, looking for components
                        if (hasattr(selection.entity, "objectType") and selection.entity.objectType == adsk.fusion.Occurrence.classType()):
                            occs.append(selection.entity)
                            if selection.entity.component:
                                for item in selection.entity.component.allOccurrences:
                                    occs.append(item)
                        else:
                            ui.messageBox('No components selected!\nPlease select some components.')
                            return
                else:
                    ui.messageBox('No components selected!\nPlease select some components.')
                    return
            else:
                occs = root.allOccurrences

            _total = occs.count
            print("Total Occurences: ", occs.count)

            if len(occs) == 0:
                ui.messageBox('There are no components in this design.')
                return

            fileDialog = ui.createFileDialog()
            fileDialog.isMultiSelectEnabled = False
            fileDialog.title = dialogTitle + " filename"
            fileDialog.filter = 'CSV (*.csv)'
            fileDialog.filterIndex = 0
            dialogResult = fileDialog.showSave()
            if dialogResult == adsk.core.DialogResults.DialogOK:
                filename = fileDialog.filename
            else:
                return

            # Gather information about each unique component
            bom = []
            for occ in occs:
                comp = occ.component
                occPath = occ.fullPathName
                print("comp: ", comp.name, " bRepBodies: ", comp.bRepBodies.count, " Occurences", comp.allOccurrences.count )
                if comp.name.startswith('_') and prefs[KEY_PREF_IGNORE_PREFIXED_COMP]:
                    _skipped += 1
                    print("Skipping _component: ", comp.name)
                    continue
                elif prefs[KEY_PREF_IGNORE_LINKED] and design != comp.parentDesign:
                    _skipped += 1
                    _linked += 1
                    print("Skipping linked component: ", comp.name)
                    continue
                elif not comp.bRepBodies.count and prefs[KEY_PREF_IGNORE_NO_BODIES]:
                    _noBody += 1
                    _skipped += 1
                    print("Skipping component w/o body: ", comp.name)
                    continue
                elif not occ.isVisible and prefs[KEY_PREF_IGNORE_INVISIBLE] is False:
                    _notVisible += 1
                    _skipped += 1
                    print("Skipping non-visible component: ", comp.name)
                    continue
                else:
                    # Loop through the components we've found already and increment unit count if this component is the same
                    jj = 0
                    for bomI in bom:
                        if bomI['component'] == comp:
                            # Increment the instance count of the existing row.  This creates updates the quantity 
                            # of the component rather than creating a whole new row in the BOM
                            bomI['instances'] += 1
                            break
                        jj += 1

                    # if jj == len of the bom we've reached the end of the list without finding the component
                    # Thus, this is a new component and we need to add it to the list.
                    if jj == len(bom):
                        # Add this component to the BOM
                        _added += 1
                        print("Adding component: ", comp.name)

                        # I would **love** to know exactly what this is for and how to avoid it.
                        # this error is the whole reason I started hacking this code :/
                        bb = self.getBodiesBoundingBox(comp.bRepBodies)
                        if not bb:
                            if ui:
                                ui.messageBox('Not all Fusion modules are loaded yet, please click on the root component to load them and try again.')
                            return

                        # Add first instance of component into BOM
                        bom.append({
                            "component": comp,
                            "parent": occPath,
                            "partnumber": comp.partNumber,
                            "name": comp.name,
                            "instances": 1,
                            "volume": self.getBodiesVolume(comp.bRepBodies),
                            "boundingBox": bb,
                            "area": self.getPhysicsArea(comp.bRepBodies),
                            "mass": self.getPhysicalMass(comp.bRepBodies),
                            "density": self.getPhysicalDensity(comp.bRepBodies),
                            "material": self.getPhysicalMaterial(comp.bRepBodies),
                            "desc": comp.description
                        })
            csvStr = self.collectData(design, bom, prefs)
            output = open(filename, 'w')
            output.writelines(csvStr)
            output.close()

            # save CutList:
            if prefs[KEY_PREF_GEN_CUT_LIST] and prefs[KEY_PREF_INC_DIMENSIONS]:
            	cutListStr = self.collectCutList(design, bom, prefs)
            	output = open(filename[:len(filename) - 4] + '_cutList.txt', 'w')
            	output.write(cutListStr)
            	output.close()

            # Save last chosen options
            design.attributes.add(cmdId, "lastUsedOptions", json.dumps(prefs))
            ui.messageBox('File written to "' + filename + '"')
        except:
            if ui:
                ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))


# Called when the plugin is initially run
def run(context):
    try:
        global ui
        global cmdId
        global dialogTitle
        global cmdDesc
        global cmdRes

        #ui.messageBox('Hello addin')

        # Get the CommandDefinitions collection.
        cmdDefs = ui.commandDefinitions
        # Create a button command definition.
        bomButton = cmdDefs.addButtonDefinition(cmdId, dialogTitle, cmdDesc, cmdRes)

        # Connect to the command created event.
        commandCreated = BOMCommandCreatedEventHandler()
        bomButton.commandCreated.add(commandCreated)
        handlers.append(commandCreated)

        # Get the CREATE panel in the model workspace.
        toolbarPanel = ui.allToolbarPanels.itemById(cmdToolbarLocation)

        # Add the button to the bottom of the panel.
        buttonControl = toolbarPanel.controls.addCommand(bomButton, "", False)
        buttonControl.isVisible = True

    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))


# Called when the pluyin is stopped
def stop(context):
    try:
        global app
        global ui
        #ui.messageBox('Stop addin')

        # Clean up the UI.
        cmdDef = ui.commandDefinitions.itemById(cmdId)
        if cmdDef:
            cmdDef.deleteMe()

        toolbarPanel = ui.allToolbarPanels.itemById(cmdToolbarLocation)
        cntrl = toolbarPanel.controls.itemById(cmdId)
        if cntrl:
            cntrl.deleteMe()

    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))
