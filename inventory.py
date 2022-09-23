from pyvim.connect import SmartConnect
from pyVmomi import vim
import xlsxwriter
import getpass



def dclist(content,worksheet, cell_format):
    worksheet.write('A1', 'DCname', cell_format)
    worksheet.write('B1', 'VMfolders', cell_format)
    worksheet.write('C1', 'noofVMs', cell_format)
    worksheet.write('D1', 'noofvApps', cell_format)
    worksheet.write('E1', 'noofStandardPortgroups', cell_format)
    worksheet.write('F1', 'noofDVswitch', cell_format)
    worksheet.write('G1', 'noofDvportgroups', cell_format)
    worksheet.write('H1', 'noofClusters', cell_format)
    worksheet.write('I1', 'noofDatastores', cell_format)
    viewType = [vim.Datacenter]
    container = content.rootFolder
    recursive = True  # whether we should look into it recursively
    containerView = content.viewManager.CreateContainerView(container, viewType, recursive)
    children = containerView.view
    r = 1
    c = 0
    for child in children:
        worksheet.write(r, c, child.name)
        c += 1
        f = 0
        vm = 0
        vapp = 0
        pg = 0
        clus = 0
        ds = 0
        dvsw = 0
        for i in child.vmFolder.childEntity:
            a = str(i)
            if a.split(':')[0] == "'vim.Folder":
                f += 1
            elif a.split(':')[0] == "'vim.VirtualMachine":
                vm += 1
            elif a.split(':')[0] == "'vim.VirtualApp":
                vapp += 1
        worksheet.write(r, c, f)
        c += 1
        f=0
        worksheet.write(r, c, vm)
        c += 1
        vm = 0
        worksheet.write(r, c, vapp)
        c += 1
        vapp = 0
        for i in child.networkFolder.childEntity:
            b = str(i)
            if 'vim.Network' in b:
                if b.split(':')[0] == "'vim.Network":
                    f += 1
            else:
                a = i.childEntity
                if a:
                    for j in a:
                        j = str(j)
                        if j.split(':')[0] == "'vim.dvs.VmwareDistributedVirtualSwitch":
                            dvsw += 1
                        elif j.split(':')[0] == "'vim.dvs.DistributedVirtualPortgroup":
                            pg += 1
        worksheet.write(r, c, f)
        c += 1
        worksheet.write(r, c, dvsw)
        c += 1
        worksheet.write(r, c, pg)
        c += 1
        for i in child.hostFolder.childEntity:
            a = str(i)
            if a.split(':')[0] == "'vim.ClusterComputeResource":
                clus += 1
        worksheet.write(r, c, clus)
        c += 1
        for i in child.datastoreFolder.childEntity:
            a = i.childEntity
            if a:
                for j in a:
                    j = str(j)
                    if j.split(':')[0] == "'vim.Datastore":
                       ds += 1
        worksheet.write(r, c, ds)
        c = 0
        r += 1

def clusterlist(content,worksheet1, cell_format):
    worksheet1.write('A1', 'ClusterName', cell_format)
    worksheet1.write('B1', 'ParentDatacenter', cell_format)
    worksheet1.write('C1', 'overallStatus', cell_format)
    worksheet1.write('D1', 'totalCpuGHz', cell_format)
    worksheet1.write('E1', 'totalMemoryGB', cell_format)
    worksheet1.write('F1', 'effectiveCpuGHz', cell_format)
    worksheet1.write('G1', 'effectiveMemory', cell_format)
    worksheet1.write('H1', 'numCpuCores', cell_format)
    worksheet1.write('I1', 'numCpuThreads', cell_format)
    worksheet1.write('J1', 'cpuDemandMhz', cell_format)
    worksheet1.write('K1', 'cpuEntitledMhz', cell_format)
    worksheet1.write('L1', 'cpuReservationMhz', cell_format)
    worksheet1.write('M1', 'memDemandMB', cell_format)
    worksheet1.write('N1', 'memEntitledMB', cell_format)
    worksheet1.write('O1', 'memReservationMB', cell_format)
    worksheet1.write('P1', 'totalCpuCapacityMhz', cell_format)
    worksheet1.write('Q1', 'totalMemCapacityMB', cell_format)
    worksheet1.write('R1', 'numHosts', cell_format)
    worksheet1.write('S1', 'numEffectiveHosts', cell_format)
    worksheet1.write('T1', 'failoverLevel', cell_format)
    worksheet1.write('U1', 'currentFailoverLevel', cell_format)
    worksheet1.write('V1', 'currentEVCModeKey', cell_format)
    worksheet1.write('W1', 'currentBalance', cell_format)
    worksheet1.write('X1', 'currentCpuFailoverResourcesPercent', cell_format)
    worksheet1.write('Y1', 'currentMemoryFailoverResourcesPercent', cell_format)
    worksheet1.write('Z1', 'drsRecommendation', cell_format)
    worksheet1.write('AA1', 'drsFault', cell_format)
    worksheet1.write('AB1', 'HAenabled', cell_format)
    worksheet1.write('AC1', 'hBDatastoreCandidatePolicy', cell_format)
    worksheet1.write('AD1', 'heartbeatDatastore', cell_format)
    worksheet1.write('AE1', 'hostMonitoring', cell_format)
    worksheet1.write('AF1', 'vmComponentProtecting', cell_format)
    worksheet1.write('AG1', 'vmMonitoring', cell_format)
    worksheet1.write('AH1', 'admissionControlEnabled', cell_format)
    worksheet1.write('AI1', 'autoComputePercentages', cell_format)
    worksheet1.write('AJ1', 'cpuFailoverResourcesPercent', cell_format)
    worksheet1.write('AK1', 'failoverLevel', cell_format)
    worksheet1.write('AL1', 'memoryFailoverResourcesPercent', cell_format)
    worksheet1.write('AM1', 'resourceReductionToToleratePercent', cell_format)
    worksheet1.write('AN1', 'isolationResponse', cell_format)
    worksheet1.write('AO1', 'restartPriority', cell_format)
    worksheet1.write('AP1', 'restartPriorityTimeout', cell_format)
    worksheet1.write('AQ1', 'enableAPDTimeoutForHosts', cell_format)
    worksheet1.write('AR1', 'vmReactionOnAPDCleared', cell_format)
    worksheet1.write('AS1', 'vmStorageProtectionForAPD', cell_format)
    worksheet1.write('AT1', 'vmStorageProtectionForPDL', cell_format)
    worksheet1.write('AU1', 'vmTerminateDelayForAPDSec', cell_format)
    worksheet1.write('AV1', 'clusterSettings', cell_format)
    worksheet1.write('AW1', 'vmToolsMonitoringSettings.enabled', cell_format)
    worksheet1.write('AX1', 'vmToolsMonitoringSettings.failureInterval', cell_format)
    worksheet1.write('AY1', 'vmToolsMonitoringSettings.maxFailureWindow', cell_format)
    worksheet1.write('AZ1', 'vmToolsMonitoringSettings.maxFailures', cell_format)
    worksheet1.write('BA1', 'vmToolsMonitoringSettings.minUpTime', cell_format)
    worksheet1.write('BB1', 'vmToolsMonitoringSettings.vmMonitoring', cell_format)
    worksheet1.write('BC1', 'drsConfig.defaultVmBehavior', cell_format)
    worksheet1.write('BD1', 'drsConfig.enableVmBehaviorOverrides', cell_format)
    worksheet1.write('BE1', 'drsConfig.enabled', cell_format)
    worksheet1.write('BF1', 'drsConfig.rules', cell_format)
    worksheet1.write('BG1', 'dpmConfigInfo.defaultDpmBehavior', cell_format)
    worksheet1.write('BH1', 'dpmConfigInfo.enabled', cell_format)
    worksheet1.write('BI1', 'DRS-VM&Host groups', cell_format)
    worksheet1.write('BJ1', 'proactiveDrsConfig.enabled', cell_format)
    worksheet1.write('BK1', 'vmSwapPlacement', cell_format)
    worksheet1.write('BL1', 'spbmEnabled', cell_format)
    viewType = [vim.ComputeResource]
    container = content.rootFolder
    recursive = True  # whether we should look into it recursively
    containerView = content.viewManager.CreateContainerView(container, viewType, recursive)
    children = containerView.view
    r = 1
    c = 0
    for child in children:
        worksheet1.write(r, c, child.name)
        c += 1
        worksheet1.write(r, c, child.parent.parent.name)
        c += 1
        worksheet1.write(r, c, child.summary.overallStatus)
        c += 1
        worksheet1.write(r, c, ((child.summary.totalCpu)/1024))
        c += 1
        worksheet1.write(r, c, ((((child.summary.totalMemory)/1024)/1024)/1024))
        c += 1
        worksheet1.write(r, c, ((child.summary.effectiveCpu)/1024))
        c += 1
        worksheet1.write(r, c, ((child.summary.effectiveMemory)/1024))
        c += 1
        worksheet1.write(r, c, child.summary.numCpuCores)
        c += 1
        worksheet1.write(r, c, child.summary.numCpuThreads)
        c += 1
        worksheet1.write(r, c, child.summary.usageSummary.cpuDemandMhz)
        c += 1
        worksheet1.write(r, c, child.summary.usageSummary.cpuEntitledMhz)
        c += 1
        worksheet1.write(r, c, child.summary.usageSummary.cpuReservationMhz)
        c += 1
        worksheet1.write(r, c, child.summary.usageSummary.memDemandMB)
        c += 1
        worksheet1.write(r, c, child.summary.usageSummary.memEntitledMB)
        c += 1
        worksheet1.write(r, c, child.summary.usageSummary.memReservationMB)
        c += 1
        worksheet1.write(r, c, child.summary.usageSummary.totalCpuCapacityMhz)
        c += 1
        worksheet1.write(r, c, child.summary.usageSummary.totalMemCapacityMB)
        c += 1
        worksheet1.write(r, c, child.summary.numHosts)
        c += 1
        worksheet1.write(r, c, child.summary.numEffectiveHosts)
        c += 1
        worksheet1.write(r, c, child.configuration.dasConfig.failoverLevel)
        c += 1
        worksheet1.write(r, c, child.summary.currentFailoverLevel)
        c += 1
        worksheet1.write(r, c, child.summary.currentEVCModeKey)
        c += 1
        worksheet1.write(r, c, child.summary.currentBalance)
        c += 1
        worksheet1.write(r, c, child.summary.admissionControlInfo.currentCpuFailoverResourcesPercent)
        c += 1
        worksheet1.write(r, c, child.summary.admissionControlInfo.currentMemoryFailoverResourcesPercent)
        c += 1
        worksheet1.write(r, c, str(child.drsRecommendation))
        c += 1
        worksheet1.write(r, c, str(child.drsFault))
        c += 1
        worksheet1.write(r, c, child.configuration.dasConfig.enabled)
        c += 1
        worksheet1.write(r, c, child.configuration.dasConfig.hBDatastoreCandidatePolicy)
        c += 1
        worksheet1.write(r, c, str(child.configuration.dasConfig.heartbeatDatastore))
        c += 1
        worksheet1.write(r, c, child.configuration.dasConfig.hostMonitoring)
        c += 1
        worksheet1.write(r, c, child.configuration.dasConfig.vmComponentProtecting)
        c += 1
        worksheet1.write(r, c, child.configuration.dasConfig.vmMonitoring)
        c += 1
        worksheet1.write(r, c, child.configuration.dasConfig.admissionControlEnabled)
        c += 1
        worksheet1.write(r, c, child.configuration.dasConfig.admissionControlPolicy.autoComputePercentages)
        c += 1
        worksheet1.write(r, c, child.configuration.dasConfig.admissionControlPolicy.cpuFailoverResourcesPercent)
        c += 1
        worksheet1.write(r, c, child.configuration.dasConfig.admissionControlPolicy.failoverLevel)
        c += 1
        worksheet1.write(r, c, child.configuration.dasConfig.admissionControlPolicy.memoryFailoverResourcesPercent)
        c += 1
        worksheet1.write(r, c, child.configuration.dasConfig.admissionControlPolicy.resourceReductionToToleratePercent)
        c += 1
        worksheet1.write(r, c, child.configuration.dasConfig.defaultVmSettings.isolationResponse)
        c += 1
        worksheet1.write(r, c, child.configuration.dasConfig.defaultVmSettings.restartPriority)
        c += 1
        worksheet1.write(r, c, child.configuration.dasConfig.defaultVmSettings.restartPriorityTimeout)
        c += 1
        worksheet1.write(r, c, child.configuration.dasConfig.defaultVmSettings.vmComponentProtectionSettings.enableAPDTimeoutForHosts)
        c += 1
        worksheet1.write(r, c, child.configuration.dasConfig.defaultVmSettings.vmComponentProtectionSettings.vmReactionOnAPDCleared)
        c += 1
        worksheet1.write(r, c, child.configuration.dasConfig.defaultVmSettings.vmComponentProtectionSettings.vmStorageProtectionForAPD)
        c += 1
        worksheet1.write(r, c, child.configuration.dasConfig.defaultVmSettings.vmComponentProtectionSettings.vmStorageProtectionForPDL)
        c += 1
        worksheet1.write(r, c, child.configuration.dasConfig.defaultVmSettings.vmComponentProtectionSettings.vmTerminateDelayForAPDSec)
        c += 1
        worksheet1.write(r, c, child.configuration.dasConfig.defaultVmSettings.vmToolsMonitoringSettings.clusterSettings)
        c += 1
        worksheet1.write(r, c,
                         child.configuration.dasConfig.defaultVmSettings.vmToolsMonitoringSettings.enabled)
        c += 1
        worksheet1.write(r, c,
                         child.configuration.dasConfig.defaultVmSettings.vmToolsMonitoringSettings.failureInterval)
        c += 1
        worksheet1.write(r, c,
                         child.configuration.dasConfig.defaultVmSettings.vmToolsMonitoringSettings.maxFailureWindow)
        c += 1
        worksheet1.write(r, c,
                         child.configuration.dasConfig.defaultVmSettings.vmToolsMonitoringSettings.maxFailures)
        c += 1
        worksheet1.write(r, c,
                         child.configuration.dasConfig.defaultVmSettings.vmToolsMonitoringSettings.minUpTime)
        c += 1
        worksheet1.write(r, c,
                         child.configuration.dasConfig.defaultVmSettings.vmToolsMonitoringSettings.vmMonitoring)
        c += 1
        worksheet1.write(r, c,
                         child.configuration.drsConfig.defaultVmBehavior)
        c += 1
        worksheet1.write(r, c,
                         child.configuration.drsConfig.enableVmBehaviorOverrides)
        c += 1
        worksheet1.write(r, c,
                         child.configuration.drsConfig.enabled)
        c += 1
        worksheet1.write(r, c,
                         str(child.configuration.rule))
        c += 1
        worksheet1.write(r, c,
                         child.configurationEx.dpmConfigInfo.defaultDpmBehavior)
        c += 1
        worksheet1.write(r, c,
                         child.configurationEx.dpmConfigInfo.enabled)
        c += 1
        worksheet1.write(r, c,
                         str(child.configurationEx.group))
        c += 1
        worksheet1.write(r, c,
                            child.configurationEx.proactiveDrsConfig.enabled)
        c += 1
        worksheet1.write(r, c,
                         child.configurationEx.vmSwapPlacement)
        c += 1
        worksheet1.write(r, c,
                         child.configurationEx.spbmEnabled)
        c += 1
        c = 0
        r += 1

def hostlist(content,worksheet2, cell_format):
    worksheet2.write('A1', 'HostName', cell_format)
    worksheet2.write('B1', 'overallStatus', cell_format)
    worksheet2.write('C1', 'powerState', cell_format)
    worksheet2.write('D1', 'standbyMode', cell_format)
    worksheet2.write('E1', 'connectionState', cell_format)
    worksheet2.write('F1', 'HAstate', cell_format)
    worksheet2.write('G1', 'lockdownMode', cell_format)
    worksheet2.write('H1', 'inMaintenanceMode', cell_format)
    worksheet2.write('I1', 'inQuarantineMode', cell_format)
    worksheet2.write('J1', 'ClusterName', cell_format)
    worksheet2.write('K1', 'DatacenterName', cell_format)
    worksheet2.write('L1', 'BuildVersion', cell_format)
    worksheet2.write('M1', 'vmotionEnabled', cell_format)
    worksheet2.write('N1', 'biosVersion', cell_format)
    worksheet2.write('O1', 'cpuMhz', cell_format)
    worksheet2.write('P1', 'cpuModel', cell_format)
    worksheet2.write('Q1', 'memorySize', cell_format)
    worksheet2.write('R1', 'model', cell_format)
    worksheet2.write('S1', 'numCpuCores', cell_format)
    worksheet2.write('T1', 'numCpuPkgs', cell_format)
    worksheet2.write('U1', 'numCpuThreads', cell_format)
    worksheet2.write('V1', 'numHBAs', cell_format)
    worksheet2.write('W1', 'numNics', cell_format)
    worksheet2.write('X1', 'uuid', cell_format)
    worksheet2.write('Y1', 'vendor', cell_format)
    worksheet2.write('Z1', 'managementServerIp', cell_format)
    worksheet2.write('AA1', 'defaultGateway', cell_format)
    worksheet2.write('AB1', 'ipV6Enabled', cell_format)
    worksheet2.write('AC1', 'currentEVCModeKey', cell_format)
    worksheet2.write('AD1', 'maxEVCModeKey', cell_format)
    worksheet2.write('AE1', 'rebootRequired', cell_format)
    worksheet2.write('AF1', 'numNodes', cell_format)
    worksheet2.write('AG1', 'distributedCpuFairness', cell_format)
    worksheet2.write('AH1', 'distributedMemoryFairness', cell_format)
    worksheet2.write('AI1', 'overallCpuUsage', cell_format)
    worksheet2.write('AJ1', 'overallMemoryUsage', cell_format)
    worksheet2.write('AK1', 'uptime', cell_format)
    worksheet2.write('AL1', 'bootTime', cell_format)
    worksheet2.write('AM1', 'cpuStatusInfo', cell_format)
    worksheet2.write('AN1', 'memoryStatusInfo', cell_format)
    worksheet2.write('AO1', 'netStackInstanceRuntimeInfo', cell_format)
    worksheet2.write('AP1', 'cpuPowrManagementcurrentPolicy', cell_format)
    worksheet2.write('AQ1', 'hardwareSupport', cell_format)
    worksheet2.write('AR1', 'ntpConfig', cell_format)
    worksheet2.write('AS1', 'StoragemountInfo', cell_format)
    worksheet2.write('AT1', 'firewalldefaultpolicy', cell_format)
    worksheet2.write('AU1', 'hyperThreadingStatus', cell_format)
    worksheet2.write('AV1', 'ioFilterInfo', cell_format)
    worksheet2.write('AW1', 'Multipathing', cell_format)
    worksheet2.write('AX1', 'dnsConfigDomainName', cell_format)
    worksheet2.write('AY1', 'dnsConfigHostName', cell_format)
    worksheet2.write('AZ1', 'ipv6Route', cell_format)
    worksheet2.write('BA1', 'ipRoute', cell_format)
    worksheet2.write('BB1', 'pnicDetails', cell_format)
    worksheet2.write('BC1', 'proxySwitch', cell_format)
    worksheet2.write('BD1', 'vnicDetails', cell_format)
    worksheet2.write('BE1', 'ESXiservices', cell_format)
    viewType = [vim.HostSystem]
    container = content.rootFolder
    recursive = True  # whether we should look into it recursively
    containerView = content.viewManager.CreateContainerView(container, viewType, recursive)
    children = containerView.view
    r = 1
    c = 0
    list = []
    for child in children:
        worksheet2.write(r, c, child.name)
        c += 1
        worksheet2.write(r, c, child.summary.overallStatus)
        c += 1
        worksheet2.write(r, c, child.summary.runtime.powerState)
        c += 1
        worksheet2.write(r, c, child.summary.runtime.standbyMode)
        c += 1
        worksheet2.write(r, c, child.summary.runtime.connectionState)
        c += 1
        if child.summary.runtime.dasHostState != None:
            worksheet2.write(r, c, child.summary.runtime.dasHostState.state)
            c += 1
        elif child.summary.runtime.dasHostState == None:
            worksheet2.write(r, c, 'None')
            c +=1
        worksheet2.write(r, c, child.config.lockdownMode)
        c += 1
        worksheet2.write(r, c, child.summary.runtime.inMaintenanceMode)
        c += 1
        worksheet2.write(r, c, child.summary.runtime.inQuarantineMode)
        c += 1
        worksheet2.write(r, c, child.parent.name)
        c += 1
        worksheet2.write(r, c, child.parent.parent.parent.name)
        c += 1
        worksheet2.write(r, c, child.summary.config.product.fullName)
        c += 1
        worksheet2.write(r, c, child.summary.config.vmotionEnabled)
        c += 1
        worksheet2.write(r, c, child.hardware.biosInfo.biosVersion)
        c += 1
        worksheet2.write(r, c, child.summary.hardware.cpuMhz)
        c += 1
        worksheet2.write(r, c, child.summary.hardware.cpuModel)
        c += 1
        worksheet2.write(r, c, child.summary.hardware.memorySize)
        c += 1
        worksheet2.write(r, c, child.summary.hardware.model)
        c += 1
        worksheet2.write(r, c, child.summary.hardware.numCpuCores)
        c += 1
        worksheet2.write(r, c, child.summary.hardware.numCpuPkgs)
        c += 1
        worksheet2.write(r, c, child.summary.hardware.numCpuThreads)
        c += 1
        worksheet2.write(r, c, child.summary.hardware.numHBAs)
        c += 1
        worksheet2.write(r, c, child.summary.hardware.numNics)
        c += 1
        worksheet2.write(r, c, child.summary.hardware.uuid)
        c += 1
        worksheet2.write(r, c, child.summary.hardware.vendor)
        c += 1
        worksheet2.write(r, c, child.summary.managementServerIp)
        c += 1
        worksheet2.write(r, c, child.config.network.ipRouteConfig.defaultGateway)
        c += 1
        worksheet2.write(r, c, child.config.network.ipV6Enabled)
        c += 1
        worksheet2.write(r, c, child.summary.currentEVCModeKey)
        c += 1
        worksheet2.write(r, c, child.summary.maxEVCModeKey)
        c += 1
        worksheet2.write(r, c, child.summary.rebootRequired)
        c += 1
        worksheet2.write(r, c, child.hardware.numaInfo.numNodes)
        c += 1
        worksheet2.write(r, c, child.summary.quickStats.distributedCpuFairness)
        c += 1
        worksheet2.write(r, c, child.summary.quickStats.distributedMemoryFairness)
        c += 1
        worksheet2.write(r, c, child.summary.quickStats.overallCpuUsage)
        c += 1
        worksheet2.write(r, c, child.summary.quickStats.overallMemoryUsage)
        c += 1
        worksheet2.write(r, c, child.summary.quickStats.uptime)
        c += 1
        worksheet2.write(r, c, str(child.summary.runtime.bootTime))
        c += 1
        for i in child.summary.runtime.healthSystemRuntime.hardwareStatusInfo.cpuStatusInfo:
            list.append(['procName: '+str(i.name), 'procSummary: '+str(i.status.summary)])
        worksheet2.write(r, c, str(list))
        list.clear()
        c += 1
        for i in child.summary.runtime.healthSystemRuntime.hardwareStatusInfo.memoryStatusInfo:
            list.append(['MemoryName: ' + str(i.name), 'MemorySummary: ' + str(i.status.summary)])
        worksheet2.write(r, c, str(list))
        list.clear()
        c += 1
        for i in child.summary.runtime.networkRuntimeInfo.netStackInstanceRuntimeInfo:
            list.append(['netStackInstanceKey: ' + str(i.netStackInstanceKey), 'State: ' + str(i.state), 'vmknickeys: '+str(i.vmknicKeys), 'currentIpV6Enabled: '+str(i.currentIpV6Enabled), 'maxNumberOfConnections: '+str(i.maxNumberOfConnections)])
        worksheet2.write(r, c, str(list))
        list.clear()
        c += 1
        worksheet2.write(r, c, child.hardware.cpuPowerManagementInfo.currentPolicy)
        c += 1
        worksheet2.write(r, c, child.hardware.cpuPowerManagementInfo.hardwareSupport)
        c += 1
        worksheet2.write(r, c, str(child.config.dateTimeInfo.ntpConfig.server))
        c += 1
        for i in child.config.fileSystemVolume.mountInfo:
            list.append(['volumeName: '+str(i.volume.name), 'MountPath: '+str(i.mountInfo.path), 'AccessMode: '+str(i.mountInfo.accessMode), 'Accessible: '+str(i.mountInfo.accessible), 'Mounted: '+str(i.mountInfo.mounted)])
        worksheet2.write(r, c, str(list))
        list.clear()
        c += 1
        worksheet2.write(r, c, str({'FWdftplcy(incomingBlocked)': child.config.firewall.defaultPolicy.incomingBlocked, 'FWdftplcy(outgoingBlocked)': child.config.firewall.defaultPolicy.outgoingBlocked}))
        c += 1
        worksheet2.write(r, c, str(child.config.hyperThread.active))
        c += 1
        for i in child.config.ioFilterInfo:
            list.append(['Name: '+ str(i.name), 'Summary: ' + str(i.summary), 'Available: ' +str(i.available)])
        worksheet2.write(r, c, str(list))
        list.clear()
        c += 1
        list.append(['MultipathName: '+str(child.config.multipathState.path[0].name), 'MultipathPathstate: ' +str(child.config.multipathState.path[0].pathState)])
        worksheet2.write(r, c, str(list))
        list.clear()
        c += 1
        worksheet2.write(r, c, child.config.network.dnsConfig.domainName)
        c += 1
        worksheet2.write(r, c, child.config.network.dnsConfig.hostName)
        c += 1
        for i in child.config.network.routeTableInfo.ipv6Route:
            list.append(['vmkDeviceName: '+str(i.deviceName), 'vmkgateway: '+ str(i.gateway), 'vmknetwork: '+str(i.network), 'vmkprefixLength: '+str(i.prefixLength)])
        worksheet2.write(r, c, str(list))
        list.clear()
        c += 1
        for i in child.config.network.routeTableInfo.ipRoute:
            list.append(['vmkDeviceName: '+str(i.deviceName), 'vmkgateway: '+ str(i.gateway), 'vmknetwork: '+str(i.network), 'vmkprefixLength: '+str(i.prefixLength)])
        worksheet2.write(r, c, str(list))
        list.clear()
        c += 1
        for i in child.config.network.pnic:
            list.append(['pnicName: ' +str(i.device), 'autoNegotiateSupported: ' +str(i.autoNegotiateSupported),
                         'isduplex: ' + str(i.linkSpeed.duplex), 'pnicSpeedMB: ' +str(i.linkSpeed.speedMb), 'pnicMAC: ' +str(i.mac),
                         'wakeOnLanSupported: ' +str(i.wakeOnLanSupported)])
        worksheet2.write(r, c, str(list))
        list.clear()
        c += 1
        for i in child.config.network.proxySwitch:
            list.append(['dvsName: ' +str(i.dvsName), 'mtu: ' +str(i.mtu), 'Totalports: ' +str(i.numPorts),
                         'AvailablePorts: ' +str(i.numPortsAvailable), 'pnic: ' +str(i.pnic)])
        worksheet2.write(r, c, str(list))
        list.clear()
        c += 1
        for i in child.config.network.vnic:
            list.append(['vmkDevice: ' +str(i.device), 'netStackInstanceKey: '+str(i.spec.netStackInstanceKey), 'mtu: '+str(i.spec.mtu), 'mac: '+str(i.spec.mac), 'ipv4Address: '+str(i.spec.ip.ipAddress),
                         'subnetMask: '+str(i.spec.ip.subnetMask), 'ipV6Address: '+str(i.spec.ip.ipV6Config.ipV6Address[0].ipAddress), 'prefixLength: '+str(i.spec.ip.ipV6Config.ipV6Address[0].prefixLength)])
        worksheet2.write(r, c, str(list))
        list.clear()
        c += 1
        for i in child.config.service.service:
            list.append(['ServiceName: '+str(i.label), 'ServiceRunningStatus: '+str(i.running)])
        worksheet2.write(r, c, str(list))
        list.clear()
        c = 0
        r += 1

def resourcepoollist(content,worksheet3, cell_format):
    worksheet3.write('A1', 'ResourcepoolName', cell_format)
    worksheet3.write('B1', 'overallStatus', cell_format)
    worksheet3.write('C1', 'noofVMs', cell_format)
    worksheet3.write('D1', 'cpuexpandableReservation', cell_format)
    worksheet3.write('E1', 'cpushareslevel', cell_format)
    worksheet3.write('F1', 'cpushares', cell_format)
    worksheet3.write('G1', 'memoryexpandableReservation', cell_format)
    worksheet3.write('H1', 'memoryshareslevel', cell_format)
    worksheet3.write('I1', 'memoryshares', cell_format)
    worksheet3.write('J1', 'balloonedMemory', cell_format)
    worksheet3.write('K1', 'compressedMemory', cell_format)
    worksheet3.write('L1', 'consumedOverheadMemory', cell_format)
    worksheet3.write('M1', 'distributedCpuEntitlement', cell_format)
    worksheet3.write('N1', 'distributedMemoryEntitlement', cell_format)
    worksheet3.write('O1', 'guestMemoryUsage', cell_format)
    worksheet3.write('P1', 'hostMemoryUsage', cell_format)
    worksheet3.write('Q1', 'overallCpuDemand', cell_format)
    worksheet3.write('R1', 'overallCpuUsage', cell_format)
    worksheet3.write('S1', 'overheadMemory', cell_format)
    worksheet3.write('T1', 'privateMemory', cell_format)
    worksheet3.write('U1', 'sharedMemory', cell_format)
    worksheet3.write('V1', 'staticCpuEntitlement', cell_format)
    worksheet3.write('W1', 'staticMemoryEntitlement', cell_format)
    worksheet3.write('X1', 'swappedMemory', cell_format)
    worksheet3.write('Y1', 'cpu.maxUsage', cell_format)
    worksheet3.write('Z1', 'cpu.overallUsage', cell_format)
    worksheet3.write('AA1', 'cpu.reservationUsed', cell_format)
    worksheet3.write('AB1', 'cpu.reservationUsedForVm', cell_format)
    worksheet3.write('AC1', 'cpu.unreservedForPool', cell_format)
    worksheet3.write('AD1', 'cpu.unreservedForVm', cell_format)
    worksheet3.write('AE1', 'memory.maxUsageGB', cell_format)
    worksheet3.write('AF1', 'memory.overallUsageGB', cell_format)
    worksheet3.write('AG1', 'memory.reservationUsedGB', cell_format)
    worksheet3.write('AH1', 'memory.reservationUsedForvmGB', cell_format)
    worksheet3.write('AI1', 'memory.unreservedForPoolGB', cell_format)
    worksheet3.write('AJ1', 'memory.unreservedForVmGB', cell_format)
    viewType = [vim.ResourcePool]
    container = content.rootFolder
    recursive = True  # whether we should look into it recursively
    containerView = content.viewManager.CreateContainerView(container, viewType, recursive)
    children = containerView.view
    r = 1
    c = 0
    for child in children:
        b = str(child)
        if 'vim.ResourcePool' in b:
            worksheet3.write(r, c, child.name)
            c += 1
            worksheet3.write(r, c, child.overallStatus)
            c += 1
            count = 0
            for i in child.vm:
                count += 1
            worksheet3.write(r, c, count)
            c += 1
            worksheet3.write(r, c, child.summary.config.cpuAllocation.expandableReservation)
            c += 1
            worksheet3.write(r, c, child.summary.config.cpuAllocation.shares.level)
            c += 1
            worksheet3.write(r, c, child.summary.config.cpuAllocation.shares.shares)
            c += 1
            worksheet3.write(r, c, child.summary.config.memoryAllocation.expandableReservation)
            c += 1
            worksheet3.write(r, c, child.summary.config.memoryAllocation.shares.level)
            c += 1
            worksheet3.write(r, c, child.summary.config.memoryAllocation.shares.shares)
            c += 1
            worksheet3.write(r, c, child.summary.quickStats.balloonedMemory)
            c += 1
            worksheet3.write(r, c, child.summary.quickStats.compressedMemory)
            c += 1
            worksheet3.write(r, c, child.summary.quickStats.consumedOverheadMemory)
            c += 1
            worksheet3.write(r, c, child.summary.quickStats.distributedCpuEntitlement)
            c += 1
            worksheet3.write(r, c, child.summary.quickStats.distributedMemoryEntitlement)
            c += 1
            worksheet3.write(r, c, child.summary.quickStats.guestMemoryUsage)
            c += 1
            worksheet3.write(r, c, child.summary.quickStats.hostMemoryUsage)
            c += 1
            worksheet3.write(r, c, child.summary.quickStats.overallCpuDemand)
            c += 1
            worksheet3.write(r, c, child.summary.quickStats.overallCpuUsage)
            c += 1
            worksheet3.write(r, c, child.summary.quickStats.overheadMemory)
            c += 1
            worksheet3.write(r, c, child.summary.quickStats.privateMemory)
            c += 1
            worksheet3.write(r, c, child.summary.quickStats.sharedMemory)
            c += 1
            worksheet3.write(r, c, child.summary.quickStats.staticCpuEntitlement)
            c += 1
            worksheet3.write(r, c, child.summary.quickStats.staticMemoryEntitlement)
            c += 1
            worksheet3.write(r, c, child.summary.quickStats.swappedMemory)
            c += 1
            worksheet3.write(r, c, child.summary.runtime.cpu.maxUsage)
            c += 1
            worksheet3.write(r, c, child.summary.runtime.cpu.overallUsage)
            c += 1
            worksheet3.write(r, c, child.summary.runtime.cpu.reservationUsed)
            c += 1
            worksheet3.write(r, c, child.summary.runtime.cpu.reservationUsedForVm)
            c += 1
            worksheet3.write(r, c, child.summary.runtime.cpu.unreservedForPool)
            c += 1
            worksheet3.write(r, c, child.summary.runtime.cpu.unreservedForVm)
            c += 1
            worksheet3.write(r, c, ((((child.summary.runtime.memory.maxUsage)/1024)/1024)/1024))
            c += 1
            worksheet3.write(r, c, ((((child.summary.runtime.memory.overallUsage)/1024)/1024)/1024))
            c += 1
            worksheet3.write(r, c, ((((child.summary.runtime.memory.reservationUsed)/1024)/1024)/1024))
            c += 1
            worksheet3.write(r, c, ((((child.summary.runtime.memory.reservationUsedForVm)/1024)/1024)/1024))
            c += 1
            worksheet3.write(r, c, ((((child.summary.runtime.memory.unreservedForPool)/1024)/1024)/1024))
            c += 1
            worksheet3.write(r, c, ((((child.summary.runtime.memory.unreservedForVm)/1024)/1024)/1024))
            c = 0
            r += 1

def vmlist(content, worksheet4, cell_format):
    worksheet4.write('A1', 'VMName', cell_format)
    worksheet4.write('B1', 'PowerState', cell_format)
    worksheet4.write('C1', 'overallStatus', cell_format)
    worksheet4.write('D1', 'vcpus', cell_format)
    worksheet4.write('E1', 'numCoresPerSocket', cell_format)
    worksheet4.write('F1', 'MemoryGB', cell_format)
    worksheet4.write('G1', 'ESXiHost', cell_format)
    worksheet4.write('H1', 'VMDKs', cell_format)
    worksheet4.write('I1', 'vNICs', cell_format)
    worksheet4.write('J1', 'IP_Address', cell_format)
    worksheet4.write('K1', 'Hardware_Version', cell_format)
    worksheet4.write('L1', 'ResourcePool', cell_format)
    worksheet4.write('M1', 'VMtools_Status', cell_format)
    worksheet4.write('N1', 'VMtools_version', cell_format)
    worksheet4.write('O1', 'VMtools_version1state', cell_format)
    worksheet4.write('P1', 'VMtools_version2state', cell_format)
    worksheet4.write('Q1', 'guestFullName', cell_format)
    worksheet4.write('R1', 'HAprotected', cell_format)
    worksheet4.write('S1', 'snapshots', cell_format)
    worksheet4.write('T1', 'cpuHotAddEnabled', cell_format)
    worksheet4.write('U1', 'memoryHotAddEnabled', cell_format)
    worksheet4.write('V1', 'guestMemoryUsage', cell_format)
    worksheet4.write('W1', 'overallCpuUsage', cell_format)
    worksheet4.write('X1', 'maxCpuUsage', cell_format)
    worksheet4.write('Y1', 'maxMemoryUsage', cell_format)
    worksheet4.write('Z1', 'cpualloc_limit', cell_format)
    worksheet4.write('AA1', 'cpuReservation', cell_format)
    worksheet4.write('AB1', 'cpualloc_share_level', cell_format)
    worksheet4.write('AC1', 'cpualloc_shares', cell_format)
    worksheet4.write('AD1', 'memalloc_limit', cell_format)
    worksheet4.write('AE1', 'memoryReservation', cell_format)
    worksheet4.write('AF1', 'memalloc_share_level', cell_format)
    worksheet4.write('AG1', 'memalloc_shares', cell_format)
    worksheet4.write('AH1', 'balloonedMemory', cell_format)
    worksheet4.write('AI1', 'consumedOverheadMemory', cell_format)
    worksheet4.write('AJ1', 'swappedMemory', cell_format)
    worksheet4.write('AK1', 'swapFile', cell_format)
    worksheet4.write('AL1', 'guestHeartbeatStatus', cell_format)
    worksheet4.write('AM1', 'uptimeSeconds', cell_format)
    worksheet4.write('AN1', 'minRequiredEVCModeKey', cell_format)
    worksheet4.write('AO1', 'faultToleranceState', cell_format)
    worksheet4.write('AP1', 'changeTrackingEnabled', cell_format)
    worksheet4.write('AQ1', 'toolsInstallerMounted', cell_format)
    worksheet4.write('AR1', 'vFlashCacheAllocation', cell_format)
    worksheet4.write('AS1', 'installBootRequired', cell_format)
    worksheet4.write('AT1', 'uuid', cell_format)
    worksheet4.write('AU1', 'vmPathName', cell_format)
    worksheet4.write('AV1', 'MksConnections', cell_format)
    viewType = [vim.VirtualMachine]
    container = content.rootFolder
    recursive = True  # whether we should look into it recursively
    containerView = content.viewManager.CreateContainerView(container, viewType, recursive)
    children = containerView.view
    r = 1
    c = 0
    for child in children:
        cpualloc = child.resourceConfig.cpuAllocation
        memalloc = child.resourceConfig.memoryAllocation
        configure = child.summary
        guest = child.guest
        conf1 = child.config
        worksheet4.write(r, c, configure.config.name)
        c += 1
        worksheet4.write(r, c, configure.runtime.powerState)
        c += 1
        worksheet4.write(r, c, child.overallStatus)
        c += 1
        worksheet4.write(r, c, configure.config.numCpu)
        c += 1
        worksheet4.write(r, c, child.config.hardware.numCoresPerSocket)
        c += 1
        worksheet4.write(r, c, ((configure.config.memorySizeMB)/1024))
        c += 1
        worksheet4.write(r, c, str(child.summary.runtime.host.name))
        c += 1
        worksheet4.write(r, c, configure.config.numVirtualDisks)
        c += 1
        worksheet4.write(r, c, configure.config.numEthernetCards)
        c += 1
        worksheet4.write(r, c, configure.guest.ipAddress)
        c += 1
        worksheet4.write(r, c, conf1.version)
        c += 1
        worksheet4.write(r, c, child.resourcePool.summary.name)
        c += 1
        worksheet4.write(r, c, configure.guest.toolsStatus)
        c += 1
        worksheet4.write(r, c, guest.toolsVersion)
        c += 1
        worksheet4.write(r, c, guest.toolsVersionStatus)
        c += 1
        worksheet4.write(r, c, guest.toolsVersionStatus2)
        c += 1
        worksheet4.write(r, c, configure.guest.guestFullName)
        c += 1
        HAprotected = configure.runtime.dasVmProtection
        if HAprotected != None:
            HAprotected = HAprotected.dasProtected
        worksheet4.write(r, c, HAprotected)
        c += 1
        snapshots = child.snapshot
        if snapshots != None:
            snapshots = child.rootSnapshot
        worksheet4.write(r, c, str(snapshots))
        c += 1
        worksheet4.write(r, c, conf1.cpuHotAddEnabled)
        c += 1
        worksheet4.write(r, c, conf1.memoryHotAddEnabled)
        c += 1
        worksheet4.write(r, c, ((configure.quickStats.guestMemoryUsage)/1024))
        c += 1
        worksheet4.write(r, c, ((configure.quickStats.overallCpuUsage)/1024))
        c += 1
        if configure.runtime.maxCpuUsage != None:
            worksheet4.write(r, c, ((configure.runtime.maxCpuUsage)/1024))
            c += 1
        elif configure.runtime.maxCpuUsage == None:
            worksheet4.write(r, c, configure.runtime.maxCpuUsage)
            c += 1
        if configure.runtime.maxCpuUsage != None:
            worksheet4.write(r, c, ((configure.runtime.maxMemoryUsage)/1024))
            c += 1
        elif configure.runtime.maxCpuUsage == None:
            worksheet4.write(r, c, configure.runtime.maxMemoryUsage)
            c += 1
        worksheet4.write(r, c, cpualloc.limit)
        c += 1
        worksheet4.write(r, c, configure.config.cpuReservation)
        c += 1
        worksheet4.write(r, c, cpualloc.shares.level)
        c += 1
        worksheet4.write(r, c, cpualloc.shares.shares)
        c += 1
        worksheet4.write(r, c, memalloc.limit)
        c += 1
        worksheet4.write(r, c, configure.config.memoryReservation)
        c += 1
        worksheet4.write(r, c, memalloc.shares.level)
        c += 1
        worksheet4.write(r, c, memalloc.shares.shares)
        c += 1
        worksheet4.write(r, c, configure.quickStats.balloonedMemory)
        c += 1
        worksheet4.write(r, c, configure.quickStats.consumedOverheadMemory)
        c += 1
        worksheet4.write(r, c, configure.quickStats.swappedMemory)
        c += 1
        worksheet4.write(r, c, child.layout.swapFile)
        c += 1
        worksheet4.write(r, c, configure.quickStats.guestHeartbeatStatus)
        c += 1
        worksheet4.write(r, c, configure.quickStats.uptimeSeconds)
        c += 1
        worksheet4.write(r, c, configure.runtime.minRequiredEVCModeKey)
        c += 1
        worksheet4.write(r, c, configure.runtime.faultToleranceState)
        c += 1
        worksheet4.write(r, c, conf1.changeTrackingEnabled)
        c += 1
        worksheet4.write(r, c, configure.runtime.toolsInstallerMounted)
        c += 1
        worksheet4.write(r, c, configure.runtime.vFlashCacheAllocation)
        c += 1
        worksheet4.write(r, c, configure.config.installBootRequired)
        c += 1
        worksheet4.write(r, c, configure.config.uuid)
        c += 1
        worksheet4.write(r, c, configure.config.vmPathName)
        c += 1
        worksheet4.write(r, c, configure.runtime.numMksConnections)
        c = 0
        r += 1

def datastoreslist(content,worksheet5, cell_format):
    worksheet5.write('A1', 'DSName', cell_format)
    worksheet5.write('B1', 'overallStatus', cell_format)
    worksheet5.write('C1', 'accessible', cell_format)
    worksheet5.write('D1', 'HostsConnected', cell_format)
    worksheet5.write('E1', 'capacity', cell_format)
    worksheet5.write('F1', 'freeSpace', cell_format)
    worksheet5.write('G1', 'maintenanceMode', cell_format)
    worksheet5.write('H1', 'multipleHostAccess', cell_format)
    worksheet5.write('I1', 'type', cell_format)
    worksheet5.write('J1', 'url', cell_format)
    worksheet5.write('K1', 'remoteHost', cell_format)
    worksheet5.write('L1', 'remotePath', cell_format)
    worksheet5.write('M1', 'congestionThreshold', cell_format)
    worksheet5.write('N1', 'congestionThresholdMode', cell_format)
    worksheet5.write('O1', 'iormConfiguration.enabled', cell_format)
    worksheet5.write('P1', 'percentOfPeakThroughput', cell_format)
    worksheet5.write('Q1', 'reservationEnabled', cell_format)
    worksheet5.write('R1', 'statsAggregationDisabled', cell_format)
    worksheet5.write('S1', 'statsCollectionEnabled', cell_format)
    viewType = [vim.Datastore]
    container = content.rootFolder
    recursive = True  # whether we should look into it recursively
    containerView = content.viewManager.CreateContainerView(container, viewType, recursive)
    children = containerView.view
    r = 1
    c = 0
    list = []
    for child in children:
        try:
            if child.info.nas:
                worksheet5.write(r, c, child.name)
                c += 1
                worksheet5.write(r, c, child.overallStatus)
                c += 1
                worksheet5.write(r, c, child.summary.accessible)
                c += 1
                for i in child.host:
                    list.append([str(i.key.name)])
                worksheet5.write(r, c, str(list))
                list.clear()
                c += 1
                worksheet5.write(r, c, child.summary.capacity)
                c += 1
                worksheet5.write(r, c, child.summary.freeSpace)
                c += 1
                worksheet5.write(r, c, child.summary.maintenanceMode)
                c += 1
                worksheet5.write(r, c, child.summary.multipleHostAccess)
                c += 1
                worksheet5.write(r, c, child.summary.type)
                c += 1
                worksheet5.write(r, c, child.summary.url)
                c += 1
                worksheet5.write(r, c, child.info.nas.remoteHost)
                c += 1
                worksheet5.write(r, c, child.info.nas.remotePath)
                c += 1
                worksheet5.write(r, c, child.iormConfiguration.congestionThreshold)
                c += 1
                worksheet5.write(r, c, child.iormConfiguration.congestionThresholdMode)
                c += 1
                worksheet5.write(r, c, child.iormConfiguration.enabled)
                c += 1
                worksheet5.write(r, c, child.iormConfiguration.percentOfPeakThroughput)
                c += 1
                worksheet5.write(r, c, child.iormConfiguration.reservationEnabled)
                c += 1
                worksheet5.write(r, c, child.iormConfiguration.statsAggregationDisabled)
                c += 1
                worksheet5.write(r, c, child.iormConfiguration.statsCollectionEnabled)
                c = 0
                r += 1
        except:
            pass

def dvslist(content,worksheet6, cell_format):
    worksheet6.write('A1', 'DVSwitchName', cell_format)
    worksheet6.write('B1', 'overallStatus', cell_format)
    worksheet6.write('C1', 'configStatus', cell_format)
    worksheet6.write('D1', 'noofPortGroup', cell_format)
    worksheet6.write('E1', 'uuid', cell_format)
    worksheet6.write('F1', 'DVSwitchversion', cell_format)
    worksheet6.write('G1', 'DVSwitchforwardingClass', cell_format)
    worksheet6.write('H1', 'numHosts', cell_format)
    worksheet6.write('I1', 'numPorts', cell_format)
    worksheet6.write('J1', 'maxMtu', cell_format)
    worksheet6.write('K1', 'multicastFilteringMode', cell_format)
    worksheet6.write('L1', 'networkResourceControlVersion', cell_format)
    worksheet6.write('M1', 'networkResourceManagementEnabled', cell_format)
    worksheet6.write('N1', 'DVPortgroupUplinkName', cell_format)
    worksheet6.write('O1', 'VMsConnected', cell_format)
    worksheet6.write('P1', 'HostsConnected', cell_format)
    worksheet6.write('Q1', 'linkDiscoveryProtocol.operation', cell_format)
    worksheet6.write('R1', 'linkDiscoveryProtocol.protocol', cell_format)
    worksheet6.write('S1', 'defaultPort.blocked.inherited', cell_format)
    worksheet6.write('T1', 'defaultPort.blocked.value', cell_format)
    worksheet6.write('U1', 'filterPolicy.filterConfig', cell_format)
    worksheet6.write('V1', 'filterPolicy.inherited', cell_format)
    worksheet6.write('W1', 'inShapingPolicy.averageBandwidth.inherited', cell_format)
    worksheet6.write('X1', 'inShapingPolicy.averageBandwidth.value', cell_format)
    worksheet6.write('Y1', 'inShapingPolicy.burstSize.inherited', cell_format)
    worksheet6.write('Z1', 'inShapingPolicy.burstSize.value', cell_format)
    worksheet6.write('AA1', 'inShapingPolicy.enabled.inherited', cell_format)
    worksheet6.write('AB1', 'inShapingPolicy.enabled.value', cell_format)
    worksheet6.write('AC1', 'inShapingPolicy.peakBandwidth.inherited', cell_format)
    worksheet6.write('AD1', 'inShapingPolicy.peakBandwidth.value', cell_format)
    worksheet6.write('AE1', 'ipfixEnabled.value', cell_format)
    worksheet6.write('AF1', 'ipfixEnabled.inherited', cell_format)
    worksheet6.write('AG1', 'lacpPolicy.enable.value', cell_format)
    worksheet6.write('AH1', 'lacpPolicy.enable.inherited', cell_format)
    worksheet6.write('AI1', 'lacpPolicy.mode.inherited', cell_format)
    worksheet6.write('AJ1', 'lacpPolicy.mode.value', cell_format)
    worksheet6.write('AK1', 'networkResourcePoolKey.inherited', cell_format)
    worksheet6.write('AL1', 'networkResourcePoolKey.value', cell_format)
    worksheet6.write('AM1', 'outShapingPolicy.averageBandwidth.inherited', cell_format)
    worksheet6.write('AN1', 'outShapingPolicy.averageBandwidth.value', cell_format)
    worksheet6.write('AO1', 'outShapingPolicy.burstSize.inherited', cell_format)
    worksheet6.write('AP1', 'outShapingPolicy.burstSize.value', cell_format)
    worksheet6.write('AQ1', 'outShapingPolicy.enabled.inherited', cell_format)
    worksheet6.write('AR1', 'outShapingPolicy.enabled.value', cell_format)
    worksheet6.write('AS1', 'outShapingPolicy.peakBandwidth.inherited', cell_format)
    worksheet6.write('AT1', 'outShapingPolicy.peakBandwidth.value', cell_format)
    worksheet6.write('AU1', 'qosTag.inherited', cell_format)
    worksheet6.write('AV1', 'qosTag.value', cell_format)
    worksheet6.write('AW1', 'allowPromiscuous.inherited', cell_format)
    worksheet6.write('AX1', 'allowPromiscuous.value', cell_format)
    worksheet6.write('AY1', 'forgedTransmits.inherited', cell_format)
    worksheet6.write('AZ1', 'forgedTransmits.value', cell_format)
    worksheet6.write('BA1', 'macChanges.inherited', cell_format)
    worksheet6.write('BB1', 'macChanges.value', cell_format)
    worksheet6.write('BC1', 'securityPolicy.inherited', cell_format)
    worksheet6.write('BD1', 'txUplink.inherited', cell_format)
    worksheet6.write('BE1', 'txUplink.value', cell_format)
    worksheet6.write('BF1', 'checkBeacon.inherited', cell_format)
    worksheet6.write('BG1', 'checkBeacon.value', cell_format)
    worksheet6.write('BH1', 'checkDuplex.inherited', cell_format)
    worksheet6.write('BI1', 'checkDuplex.value', cell_format)
    worksheet6.write('BJ1', 'fullDuplex.inherited', cell_format)
    worksheet6.write('BK1', 'fullDuplex.value', cell_format)
    worksheet6.write('BL1', 'checkErrorPercent.inherited', cell_format)
    worksheet6.write('BM1', 'checkErrorPercent.value', cell_format)
    worksheet6.write('BN1', 'checkSpeed.inherited', cell_format)
    worksheet6.write('BO1', 'checkSpeed.value', cell_format)
    worksheet6.write('BP1', 'percentage.inherited', cell_format)
    worksheet6.write('BQ1', 'percentage.value', cell_format)
    worksheet6.write('BR1', 'speed.inherited', cell_format)
    worksheet6.write('BS1', 'speed.value', cell_format)
    worksheet6.write('BT1', 'failureCriteria.inherited', cell_format)
    worksheet6.write('BU1', 'uplinkTeamingPolicy.inherited', cell_format)
    worksheet6.write('BV1', 'notifySwitches.inherited', cell_format)
    worksheet6.write('BW1', 'notifySwitches.value', cell_format)
    worksheet6.write('BX1', 'uplinkTeamingPolicy.inherited', cell_format)
    worksheet6.write('BY1', 'uplinkTeamingPolicy.value', cell_format)
    worksheet6.write('BZ1', 'uplinkTeamingPolicy.reversePolicy.inherited', cell_format)
    worksheet6.write('CA1', 'uplinkTeamingPolicy.reversePolicy.value', cell_format)
    worksheet6.write('CB1', 'uplinkTeamingPolicy.rollingOrder.inherited', cell_format)
    worksheet6.write('CC1', 'uplinkTeamingPolicy.rollingOrder.value', cell_format)
    worksheet6.write('CD1', 'activeUplinkPort', cell_format)
    worksheet6.write('CE1', 'uplinkPortOrder.inherited', cell_format)
    worksheet6.write('CF1', 'standbyUplinkPort', cell_format)
    worksheet6.write('CG1', 'vlan.inherited', cell_format)
    worksheet6.write('CH1', 'vlan.vlanId', cell_format)
    viewType = [vim.DistributedVirtualSwitch]
    container = content.rootFolder
    recursive = True  # whether we should look into it recursively
    containerView = content.viewManager.CreateContainerView(container, viewType, recursive)
    children = containerView.view
    r = 1
    c = 0
    list = []
    count = 0
    for child in children:
        worksheet6.write(r, c, child.name)
        c += 1
        worksheet6.write(r, c, child.overallStatus)
        c += 1
        worksheet6.write(r, c, child.configStatus)
        c += 1
        for i in child.portgroup:
            count += 1
        worksheet6.write(r, c, count)
        c += 1
        worksheet6.write(r, c, child.uuid)
        c += 1
        worksheet6.write(r, c, child.summary.productInfo.version)
        c += 1
        worksheet6.write(r, c, child.summary.productInfo.forwardingClass)
        c += 1
        worksheet6.write(r, c, child.summary.numHosts)
        c += 1
        worksheet6.write(r, c, child.summary.numPorts)
        c += 1
        worksheet6.write(r, c, str(child.config.maxMtu))
        c += 1
        worksheet6.write(r, c, str(child.config.multicastFilteringMode))
        c += 1
        worksheet6.write(r, c, str(child.config.networkResourceControlVersion))
        c += 1
        worksheet6.write(r, c, str(child.config.networkResourceManagementEnabled))
        c += 1
        worksheet6.write(r, c, str(child.config.uplinkPortgroup[0].name))
        c += 1
        worksheet6.write(r, c, str(child.summary.vm))
        c += 1
        for i in child.runtime.hostMemberRuntime:
            list.append(['HostName: '+str(i.host.name), 'HostConnectionstatus: '+str(i.status), 'HostStatusDetail: '+str(i.statusDetail)])
        worksheet6.write(r, c, str(list))
        list.clear()
        c += 1
        worksheet6.write(r, c, child.config.linkDiscoveryProtocolConfig.operation)
        c += 1
        worksheet6.write(r, c, child.config.linkDiscoveryProtocolConfig.protocol)
        c += 1
        worksheet6.write(r, c, child.config.defaultPortConfig.blocked.inherited)
        c += 1
        worksheet6.write(r, c, child.config.defaultPortConfig.blocked.value)
        c += 1
        worksheet6.write(r, c, str(child.config.defaultPortConfig.filterPolicy.filterConfig))
        c += 1
        worksheet6.write(r, c, child.config.defaultPortConfig.filterPolicy.inherited)
        c += 1
        worksheet6.write(r, c, child.config.defaultPortConfig.inShapingPolicy.averageBandwidth.inherited)
        c += 1
        worksheet6.write(r, c, child.config.defaultPortConfig.inShapingPolicy.averageBandwidth.value)
        c += 1
        worksheet6.write(r, c, child.config.defaultPortConfig.inShapingPolicy.burstSize.inherited)
        c += 1
        worksheet6.write(r, c, child.config.defaultPortConfig.inShapingPolicy.burstSize.value)
        c += 1
        worksheet6.write(r, c, child.config.defaultPortConfig.inShapingPolicy.enabled.inherited)
        c += 1
        worksheet6.write(r, c, child.config.defaultPortConfig.inShapingPolicy.enabled.value)
        c += 1
        worksheet6.write(r, c, child.config.defaultPortConfig.inShapingPolicy.peakBandwidth.inherited)
        c += 1
        worksheet6.write(r, c, child.config.defaultPortConfig.inShapingPolicy.peakBandwidth.value)
        c += 1
        worksheet6.write(r, c, child.config.defaultPortConfig.ipfixEnabled.value)
        c += 1
        worksheet6.write(r, c, child.config.defaultPortConfig.ipfixEnabled.inherited)
        c += 1
        worksheet6.write(r, c, child.config.defaultPortConfig.lacpPolicy.enable.value)
        c += 1
        worksheet6.write(r, c, child.config.defaultPortConfig.lacpPolicy.enable.inherited)
        c += 1
        worksheet6.write(r, c, child.config.defaultPortConfig.lacpPolicy.mode.inherited)
        c += 1
        worksheet6.write(r, c, child.config.defaultPortConfig.lacpPolicy.mode.value)
        c += 1
        worksheet6.write(r, c, child.config.defaultPortConfig.networkResourcePoolKey.inherited)
        c += 1
        worksheet6.write(r, c, child.config.defaultPortConfig.networkResourcePoolKey.value)
        c += 1
        worksheet6.write(r, c, child.config.defaultPortConfig.outShapingPolicy.averageBandwidth.inherited)
        c += 1
        worksheet6.write(r, c, child.config.defaultPortConfig.outShapingPolicy.averageBandwidth.value)
        c += 1
        worksheet6.write(r, c, child.config.defaultPortConfig.outShapingPolicy.burstSize.inherited)
        c += 1
        worksheet6.write(r, c, child.config.defaultPortConfig.outShapingPolicy.burstSize.value)
        c += 1
        worksheet6.write(r, c, child.config.defaultPortConfig.outShapingPolicy.enabled.inherited)
        c += 1
        worksheet6.write(r, c, child.config.defaultPortConfig.outShapingPolicy.enabled.value)
        c += 1
        worksheet6.write(r, c, child.config.defaultPortConfig.outShapingPolicy.peakBandwidth.inherited)
        c += 1
        worksheet6.write(r, c, child.config.defaultPortConfig.outShapingPolicy.peakBandwidth.value)
        c += 1
        worksheet6.write(r, c, child.config.defaultPortConfig.qosTag.inherited)
        c += 1
        worksheet6.write(r, c, child.config.defaultPortConfig.qosTag.value)
        c += 1
        worksheet6.write(r, c, child.config.defaultPortConfig.securityPolicy.allowPromiscuous.inherited)
        c += 1
        worksheet6.write(r, c, child.config.defaultPortConfig.securityPolicy.allowPromiscuous.value)
        c += 1
        worksheet6.write(r, c, child.config.defaultPortConfig.securityPolicy.forgedTransmits.inherited)
        c += 1
        worksheet6.write(r, c, child.config.defaultPortConfig.securityPolicy.forgedTransmits.value)
        c += 1
        worksheet6.write(r, c, child.config.defaultPortConfig.securityPolicy.macChanges.inherited)
        c += 1
        worksheet6.write(r, c, child.config.defaultPortConfig.securityPolicy.macChanges.value)
        c += 1
        worksheet6.write(r, c, child.config.defaultPortConfig.securityPolicy.inherited)
        c += 1
        worksheet6.write(r, c, child.config.defaultPortConfig.txUplink.inherited)
        c += 1
        worksheet6.write(r, c, child.config.defaultPortConfig.txUplink.value)
        c += 1
        worksheet6.write(r, c, child.config.defaultPortConfig.uplinkTeamingPolicy.failureCriteria.checkBeacon.inherited)
        c += 1
        worksheet6.write(r, c, child.config.defaultPortConfig.uplinkTeamingPolicy.failureCriteria.checkBeacon.value)
        c += 1
        worksheet6.write(r, c, child.config.defaultPortConfig.uplinkTeamingPolicy.failureCriteria.checkDuplex.inherited)
        c += 1
        worksheet6.write(r, c, child.config.defaultPortConfig.uplinkTeamingPolicy.failureCriteria.checkDuplex.value)
        c += 1
        worksheet6.write(r, c, child.config.defaultPortConfig.uplinkTeamingPolicy.failureCriteria.checkErrorPercent.inherited)
        c += 1
        worksheet6.write(r, c, child.config.defaultPortConfig.uplinkTeamingPolicy.failureCriteria.checkErrorPercent.value)
        c += 1
        worksheet6.write(r, c, child.config.defaultPortConfig.uplinkTeamingPolicy.failureCriteria.checkSpeed.inherited)
        c += 1
        worksheet6.write(r, c, child.config.defaultPortConfig.uplinkTeamingPolicy.failureCriteria.checkSpeed.value)
        c += 1
        worksheet6.write(r, c, child.config.defaultPortConfig.uplinkTeamingPolicy.failureCriteria.fullDuplex.inherited)
        c += 1
        worksheet6.write(r, c, child.config.defaultPortConfig.uplinkTeamingPolicy.failureCriteria.fullDuplex.value)
        c += 1
        worksheet6.write(r, c, child.config.defaultPortConfig.uplinkTeamingPolicy.failureCriteria.percentage.inherited)
        c += 1
        worksheet6.write(r, c, child.config.defaultPortConfig.uplinkTeamingPolicy.failureCriteria.percentage.value)
        c += 1
        worksheet6.write(r, c, child.config.defaultPortConfig.uplinkTeamingPolicy.failureCriteria.speed.inherited)
        c += 1
        worksheet6.write(r, c, child.config.defaultPortConfig.uplinkTeamingPolicy.failureCriteria.speed.value)
        c += 1
        worksheet6.write(r, c, child.config.defaultPortConfig.uplinkTeamingPolicy.failureCriteria.inherited)
        c += 1
        worksheet6.write(r, c, child.config.defaultPortConfig.uplinkTeamingPolicy.inherited)
        c += 1
        worksheet6.write(r, c, child.config.defaultPortConfig.uplinkTeamingPolicy.notifySwitches.inherited)
        c += 1
        worksheet6.write(r, c, child.config.defaultPortConfig.uplinkTeamingPolicy.notifySwitches.value)
        c += 1
        worksheet6.write(r, c, child.config.defaultPortConfig.uplinkTeamingPolicy.policy.inherited)
        c += 1
        worksheet6.write(r, c, child.config.defaultPortConfig.uplinkTeamingPolicy.policy.value)
        c += 1
        worksheet6.write(r, c,
                         child.config.defaultPortConfig.uplinkTeamingPolicy.reversePolicy.inherited)
        c += 1
        worksheet6.write(r, c,
                         child.config.defaultPortConfig.uplinkTeamingPolicy.reversePolicy.value)
        c += 1
        worksheet6.write(r, c, child.config.defaultPortConfig.uplinkTeamingPolicy.rollingOrder.inherited)
        c += 1
        worksheet6.write(r, c, child.config.defaultPortConfig.uplinkTeamingPolicy.rollingOrder.value)
        c += 1
        worksheet6.write(r, c,
                         str(child.config.defaultPortConfig.uplinkTeamingPolicy.uplinkPortOrder.activeUplinkPort))
        c += 1
        worksheet6.write(r, c, child.config.defaultPortConfig.uplinkTeamingPolicy.uplinkPortOrder.inherited)
        c += 1
        worksheet6.write(r, c, str(child.config.defaultPortConfig.uplinkTeamingPolicy.uplinkPortOrder.standbyUplinkPort))
        c += 1
        worksheet6.write(r, c, child.config.defaultPortConfig.vlan.inherited)
        c += 1
        worksheet6.write(r, c,
                         str(child.config.defaultPortConfig.vlan.vlanId))
        c = 0
        r += 1

def pglist(content,worksheet7, cell_format):
    worksheet7.write('A1', 'DVPGName', cell_format)
    worksheet7.write('B1', 'configStatus', cell_format)
    worksheet7.write('C1', 'key', cell_format)
    worksheet7.write('D1', 'overallStatus', cell_format)
    worksheet7.write('E1', 'DVswitchName', cell_format)
    worksheet7.write('F1', 'accessible', cell_format)
    worksheet7.write('G1', 'ipPoolId', cell_format)
    worksheet7.write('H1', 'ipPoolName', cell_format)
    worksheet7.write('I1', 'VMsConnected', cell_format)
    worksheet7.write('J1', 'autoExpand', cell_format)
    worksheet7.write('K1', 'numPorts', cell_format)
    worksheet7.write('L1', 'type', cell_format)
    worksheet7.write('M1', 'uplink', cell_format)
    worksheet7.write('N1', 'blockOverrideAllowed', cell_format)
    worksheet7.write('O1', 'ipfixOverrideAllowed', cell_format)
    worksheet7.write('P1', 'livePortMovingAllowed', cell_format)
    worksheet7.write('Q1', 'networkResourcePoolOverrideAllowed', cell_format)
    worksheet7.write('R1', 'portConfigResetAtDisconnect', cell_format)
    worksheet7.write('S1', 'securityPolicyOverrideAllowed', cell_format)
    worksheet7.write('T1', 'shapingOverrideAllowed', cell_format)
    worksheet7.write('U1', 'trafficFilterOverrideAllowed', cell_format)
    worksheet7.write('V1', 'uplinkTeamingOverrideAllowed', cell_format)
    worksheet7.write('W1', 'vendorConfigOverrideAllowed', cell_format)
    worksheet7.write('X1', 'vlanOverrideAllowed', cell_format)
    worksheet7.write('Y1', 'defaultPort.blocked.inherited', cell_format)
    worksheet7.write('Z1', 'defaultPort.blocked.value', cell_format)
    worksheet7.write('AA1', 'filterPolicy.filterConfig', cell_format)
    worksheet7.write('AB1', 'filterPolicy.inherited', cell_format)
    worksheet7.write('AC1', 'inShapingPolicy.averageBandwidth.inherited', cell_format)
    worksheet7.write('AD1', 'inShapingPolicy.averageBandwidth.value', cell_format)
    worksheet7.write('AE1', 'inShapingPolicy.burstSize.inherited', cell_format)
    worksheet7.write('AF1', 'inShapingPolicy.burstSize.value', cell_format)
    worksheet7.write('AG1', 'inShapingPolicy.enabled.inherited', cell_format)
    worksheet7.write('AH1', 'inShapingPolicy.enabled.value', cell_format)
    worksheet7.write('AI1', 'inShapingPolicy.peakBandwidth.inherited', cell_format)
    worksheet7.write('AJ1', 'inShapingPolicy.peakBandwidth.value', cell_format)
    worksheet7.write('AK1', 'ipfixEnabled.value', cell_format)
    worksheet7.write('AL1', 'ipfixEnabled.inherited', cell_format)
    worksheet7.write('AM1', 'lacpPolicy.enable.value', cell_format)
    worksheet7.write('AN1', 'lacpPolicy.enable.inherited', cell_format)
    worksheet7.write('AO1', 'lacpPolicy.mode.inherited', cell_format)
    worksheet7.write('AP1', 'lacpPolicy.mode.value', cell_format)
    worksheet7.write('AQ1', 'networkResourcePoolKey.inherited', cell_format)
    worksheet7.write('AR1', 'networkResourcePoolKey.value', cell_format)
    worksheet7.write('AS1', 'outShapingPolicy.averageBandwidth.inherited', cell_format)
    worksheet7.write('AT1', 'outShapingPolicy.averageBandwidth.value', cell_format)
    worksheet7.write('AU1', 'outShapingPolicy.burstSize.inherited', cell_format)
    worksheet7.write('AV1', 'outShapingPolicy.burstSize.value', cell_format)
    worksheet7.write('AW1', 'outShapingPolicy.enabled.inherited', cell_format)
    worksheet7.write('AX1', 'outShapingPolicy.enabled.value', cell_format)
    worksheet7.write('AY1', 'outShapingPolicy.peakBandwidth.inherited', cell_format)
    worksheet7.write('AZ1', 'outShapingPolicy.peakBandwidth.value', cell_format)
    worksheet7.write('BA1', 'qosTag.inherited', cell_format)
    worksheet7.write('BB1', 'qosTag.value', cell_format)
    worksheet7.write('BC1', 'allowPromiscuous.inherited', cell_format)
    worksheet7.write('BD1', 'allowPromiscuous.value', cell_format)
    worksheet7.write('BE1', 'forgedTransmits.inherited', cell_format)
    worksheet7.write('BF1', 'forgedTransmits.value', cell_format)
    worksheet7.write('BG1', 'macChanges.inherited', cell_format)
    worksheet7.write('BH1', 'macChanges.value', cell_format)
    worksheet7.write('BI1', 'securityPolicy.inherited', cell_format)
    worksheet7.write('BJ1', 'txUplink.inherited', cell_format)
    worksheet7.write('BK1', 'txUplink.value', cell_format)
    worksheet7.write('BL1', 'checkBeacon.inherited', cell_format)
    worksheet7.write('BM1', 'checkBeacon.value', cell_format)
    worksheet7.write('BN1', 'checkDuplex.inherited', cell_format)
    worksheet7.write('BO1', 'checkDuplex.value', cell_format)
    worksheet7.write('BP1', 'fullDuplex.inherited', cell_format)
    worksheet7.write('BQ1', 'fullDuplex.value', cell_format)
    worksheet7.write('BR1', 'checkErrorPercent.inherited', cell_format)
    worksheet7.write('BS1', 'checkErrorPercent.value', cell_format)
    worksheet7.write('BT1', 'checkSpeed.inherited', cell_format)
    worksheet7.write('BU1', 'checkSpeed.value', cell_format)
    worksheet7.write('BV1', 'percentage.inherited', cell_format)
    worksheet7.write('BW1', 'percentage.value', cell_format)
    worksheet7.write('BX1', 'speed.inherited', cell_format)
    worksheet7.write('BY1', 'speed.value', cell_format)
    worksheet7.write('BZ1', 'failureCriteria.inherited', cell_format)
    worksheet7.write('CA1', 'uplinkTeamingPolicy.inherited', cell_format)
    worksheet7.write('CB1', 'notifySwitches.inherited', cell_format)
    worksheet7.write('CC1', 'notifySwitches.value', cell_format)
    worksheet7.write('CD1', 'uplinkTeamingPolicy.inherited', cell_format)
    worksheet7.write('CE1', 'uplinkTeamingPolicy.value', cell_format)
    worksheet7.write('CF1', 'uplinkTeamingPolicy.reversePolicy.inherited', cell_format)
    worksheet7.write('CG1', 'uplinkTeamingPolicy.reversePolicy.value', cell_format)
    worksheet7.write('CH1', 'uplinkTeamingPolicy.rollingOrder.inherited', cell_format)
    worksheet7.write('CI1', 'uplinkTeamingPolicy.rollingOrder.value', cell_format)
    worksheet7.write('CJ1', 'activeUplinkPort', cell_format)
    worksheet7.write('CK1', 'uplinkPortOrder.inherited', cell_format)
    worksheet7.write('CL1', 'standbyUplinkPort', cell_format)
    worksheet7.write('CM1', 'vlan.inherited', cell_format)
    worksheet7.write('CN1', 'vlan.vlanId', cell_format)
    viewType = [vim.dvs.DistributedVirtualPortgroup]
    container = content.rootFolder
    recursive = True  # whether we should look into it recursively
    containerView = content.viewManager.CreateContainerView(container, viewType, recursive)
    children = containerView.view
    r = 1
    c = 0
    list = []
    count = 0
    for child in children:
        worksheet7.write(r, c, child.name)
        c += 1
        worksheet7.write(r, c, child.configStatus)
        c += 1
        worksheet7.write(r, c, child.key)
        c += 1
        worksheet7.write(r, c, child.overallStatus)
        c += 1
        worksheet7.write(r, c, child.parent.name)
        c += 1
        worksheet7.write(r, c, child.summary.accessible)
        c += 1
        worksheet7.write(r, c, child.summary.ipPoolId)
        c += 1
        worksheet7.write(r, c, child.summary.ipPoolName)
        c += 1
        worksheet7.write(r, c, str(child.vm))
        c += 1
        worksheet7.write(r, c, child.config.autoExpand)
        c += 1
        worksheet7.write(r, c, child.config.numPorts)
        c += 1
        worksheet7.write(r, c, child.config.type)
        c += 1
        worksheet7.write(r, c, child.config.uplink)
        c += 1
        worksheet7.write(r, c, child.config.policy.blockOverrideAllowed)
        c += 1
        worksheet7.write(r, c, child.config.policy.ipfixOverrideAllowed)
        c += 1
        worksheet7.write(r, c, child.config.policy.livePortMovingAllowed)
        c += 1
        worksheet7.write(r, c, child.config.policy.networkResourcePoolOverrideAllowed)
        c += 1
        worksheet7.write(r, c, child.config.policy.portConfigResetAtDisconnect)
        c += 1
        worksheet7.write(r, c, child.config.policy.securityPolicyOverrideAllowed)
        c += 1
        worksheet7.write(r, c, child.config.policy.shapingOverrideAllowed)
        c += 1
        worksheet7.write(r, c, child.config.policy.trafficFilterOverrideAllowed)
        c += 1
        worksheet7.write(r, c, child.config.policy.uplinkTeamingOverrideAllowed)
        c += 1
        worksheet7.write(r, c, child.config.policy.vendorConfigOverrideAllowed)
        c += 1
        worksheet7.write(r, c, child.config.policy.vlanOverrideAllowed)
        c += 1
        worksheet7.write(r, c, child.config.defaultPortConfig.blocked.inherited)
        c += 1
        worksheet7.write(r, c, child.config.defaultPortConfig.blocked.value)
        c += 1
        worksheet7.write(r, c, str(child.config.defaultPortConfig.filterPolicy.filterConfig))
        c += 1
        worksheet7.write(r, c, child.config.defaultPortConfig.filterPolicy.inherited)
        c += 1
        worksheet7.write(r, c, child.config.defaultPortConfig.inShapingPolicy.averageBandwidth.inherited)
        c += 1
        worksheet7.write(r, c, child.config.defaultPortConfig.inShapingPolicy.averageBandwidth.value)
        c += 1
        worksheet7.write(r, c, child.config.defaultPortConfig.inShapingPolicy.burstSize.inherited)
        c += 1
        worksheet7.write(r, c, child.config.defaultPortConfig.inShapingPolicy.burstSize.value)
        c += 1
        worksheet7.write(r, c, child.config.defaultPortConfig.inShapingPolicy.enabled.inherited)
        c += 1
        worksheet7.write(r, c, child.config.defaultPortConfig.inShapingPolicy.enabled.value)
        c += 1
        worksheet7.write(r, c, child.config.defaultPortConfig.inShapingPolicy.peakBandwidth.inherited)
        c += 1
        worksheet7.write(r, c, child.config.defaultPortConfig.inShapingPolicy.peakBandwidth.value)
        c += 1
        worksheet7.write(r, c, child.config.defaultPortConfig.ipfixEnabled.value)
        c += 1
        worksheet7.write(r, c, child.config.defaultPortConfig.ipfixEnabled.inherited)
        c += 1
        worksheet7.write(r, c, child.config.defaultPortConfig.lacpPolicy.enable.value)
        c += 1
        worksheet7.write(r, c, child.config.defaultPortConfig.lacpPolicy.enable.inherited)
        c += 1
        worksheet7.write(r, c, child.config.defaultPortConfig.lacpPolicy.mode.inherited)
        c += 1
        worksheet7.write(r, c, child.config.defaultPortConfig.lacpPolicy.mode.value)
        c += 1
        worksheet7.write(r, c, child.config.defaultPortConfig.networkResourcePoolKey.inherited)
        c += 1
        worksheet7.write(r, c, child.config.defaultPortConfig.networkResourcePoolKey.value)
        c += 1
        worksheet7.write(r, c, child.config.defaultPortConfig.outShapingPolicy.averageBandwidth.inherited)
        c += 1
        worksheet7.write(r, c, child.config.defaultPortConfig.outShapingPolicy.averageBandwidth.value)
        c += 1
        worksheet7.write(r, c, child.config.defaultPortConfig.outShapingPolicy.burstSize.inherited)
        c += 1
        worksheet7.write(r, c, child.config.defaultPortConfig.outShapingPolicy.burstSize.value)
        c += 1
        worksheet7.write(r, c, child.config.defaultPortConfig.outShapingPolicy.enabled.inherited)
        c += 1
        worksheet7.write(r, c, child.config.defaultPortConfig.outShapingPolicy.enabled.value)
        c += 1
        worksheet7.write(r, c, child.config.defaultPortConfig.outShapingPolicy.peakBandwidth.inherited)
        c += 1
        worksheet7.write(r, c, child.config.defaultPortConfig.outShapingPolicy.peakBandwidth.value)
        c += 1
        worksheet7.write(r, c, child.config.defaultPortConfig.qosTag.inherited)
        c += 1
        worksheet7.write(r, c, child.config.defaultPortConfig.qosTag.value)
        c += 1
        worksheet7.write(r, c, child.config.defaultPortConfig.securityPolicy.allowPromiscuous.inherited)
        c += 1
        worksheet7.write(r, c, child.config.defaultPortConfig.securityPolicy.allowPromiscuous.value)
        c += 1
        worksheet7.write(r, c, child.config.defaultPortConfig.securityPolicy.forgedTransmits.inherited)
        c += 1
        worksheet7.write(r, c, child.config.defaultPortConfig.securityPolicy.forgedTransmits.value)
        c += 1
        worksheet7.write(r, c, child.config.defaultPortConfig.securityPolicy.macChanges.inherited)
        c += 1
        worksheet7.write(r, c, child.config.defaultPortConfig.securityPolicy.macChanges.value)
        c += 1
        worksheet7.write(r, c, child.config.defaultPortConfig.securityPolicy.inherited)
        c += 1
        worksheet7.write(r, c, child.config.defaultPortConfig.txUplink.inherited)
        c += 1
        worksheet7.write(r, c, child.config.defaultPortConfig.txUplink.value)
        c += 1
        worksheet7.write(r, c, child.config.defaultPortConfig.uplinkTeamingPolicy.failureCriteria.checkBeacon.inherited)
        c += 1
        worksheet7.write(r, c, child.config.defaultPortConfig.uplinkTeamingPolicy.failureCriteria.checkBeacon.value)
        c += 1
        worksheet7.write(r, c, child.config.defaultPortConfig.uplinkTeamingPolicy.failureCriteria.checkDuplex.inherited)
        c += 1
        worksheet7.write(r, c, child.config.defaultPortConfig.uplinkTeamingPolicy.failureCriteria.checkDuplex.value)
        c += 1
        worksheet7.write(r, c,
                         child.config.defaultPortConfig.uplinkTeamingPolicy.failureCriteria.checkErrorPercent.inherited)
        c += 1
        worksheet7.write(r, c,
                         child.config.defaultPortConfig.uplinkTeamingPolicy.failureCriteria.checkErrorPercent.value)
        c += 1
        worksheet7.write(r, c, child.config.defaultPortConfig.uplinkTeamingPolicy.failureCriteria.checkSpeed.inherited)
        c += 1
        worksheet7.write(r, c, child.config.defaultPortConfig.uplinkTeamingPolicy.failureCriteria.checkSpeed.value)
        c += 1
        worksheet7.write(r, c, child.config.defaultPortConfig.uplinkTeamingPolicy.failureCriteria.fullDuplex.inherited)
        c += 1
        worksheet7.write(r, c, child.config.defaultPortConfig.uplinkTeamingPolicy.failureCriteria.fullDuplex.value)
        c += 1
        worksheet7.write(r, c, child.config.defaultPortConfig.uplinkTeamingPolicy.failureCriteria.percentage.inherited)
        c += 1
        worksheet7.write(r, c, child.config.defaultPortConfig.uplinkTeamingPolicy.failureCriteria.percentage.value)
        c += 1
        worksheet7.write(r, c, child.config.defaultPortConfig.uplinkTeamingPolicy.failureCriteria.speed.inherited)
        c += 1
        worksheet7.write(r, c, child.config.defaultPortConfig.uplinkTeamingPolicy.failureCriteria.speed.value)
        c += 1
        worksheet7.write(r, c, child.config.defaultPortConfig.uplinkTeamingPolicy.failureCriteria.inherited)
        c += 1
        worksheet7.write(r, c, child.config.defaultPortConfig.uplinkTeamingPolicy.inherited)
        c += 1
        worksheet7.write(r, c, child.config.defaultPortConfig.uplinkTeamingPolicy.notifySwitches.inherited)
        c += 1
        worksheet7.write(r, c, child.config.defaultPortConfig.uplinkTeamingPolicy.notifySwitches.value)
        c += 1
        worksheet7.write(r, c, child.config.defaultPortConfig.uplinkTeamingPolicy.policy.inherited)
        c += 1
        worksheet7.write(r, c, child.config.defaultPortConfig.uplinkTeamingPolicy.policy.value)
        c += 1
        worksheet7.write(r, c,
                         child.config.defaultPortConfig.uplinkTeamingPolicy.reversePolicy.inherited)
        c += 1
        worksheet7.write(r, c,
                         child.config.defaultPortConfig.uplinkTeamingPolicy.reversePolicy.value)
        c += 1
        worksheet7.write(r, c, child.config.defaultPortConfig.uplinkTeamingPolicy.rollingOrder.inherited)
        c += 1
        worksheet7.write(r, c, child.config.defaultPortConfig.uplinkTeamingPolicy.rollingOrder.value)
        c += 1
        worksheet7.write(r, c,
                         str(child.config.defaultPortConfig.uplinkTeamingPolicy.uplinkPortOrder.activeUplinkPort))
        c += 1
        worksheet7.write(r, c, child.config.defaultPortConfig.uplinkTeamingPolicy.uplinkPortOrder.inherited)
        c += 1
        worksheet7.write(r, c,
                         str(child.config.defaultPortConfig.uplinkTeamingPolicy.uplinkPortOrder.standbyUplinkPort))
        c += 1
        worksheet7.write(r, c, child.config.defaultPortConfig.vlan.inherited)
        c += 1
        worksheet7.write(r, c,
                         str(child.config.defaultPortConfig.vlan.vlanId))
    c = 0
    r += 1


if __name__ == '__main__':
    Hostname = input("Enter the vCenter HostName: ")
    vcuser = input("Enter the vCenter UserName: ")
    vcpasswd = getpass.getpass(prompt='Enter vCenter Password: ')
    service_instance = SmartConnect(host=Hostname, user=vcuser, pwd=vcpasswd)
    content = service_instance.RetrieveContent()
    if content:
        workbook = xlsxwriter.Workbook('<Output Xlsx Path>')
        worksheet = workbook.add_worksheet('Datacenters')
        worksheet1 = workbook.add_worksheet('Clusters')
        worksheet2 = workbook.add_worksheet('Host System')
        worksheet3 = workbook.add_worksheet('Resource_Pools')
        worksheet4 = workbook.add_worksheet('Virtual_Machines')
        worksheet5 = workbook.add_worksheet('Datastores')
        worksheet6 = workbook.add_worksheet('DVSwitch')
        worksheet7 = workbook.add_worksheet('Portgroup')
        worksheet8 = workbook.add_worksheet('Folder')
        cell_format = workbook.add_format({'bold': True, 'font_color': 'green'})
        print("##### Gathering Datacenter List #####")
        dclist(content,worksheet,cell_format)
        print("##### Gathering Cluster List #####")
        clusterlist(content,worksheet1,cell_format)
        print("##### Gathering HostSystem List #####")
        hostlist(content,worksheet2,cell_format)
        print("##### Gathering Resourcepool List #####")
        resourcepoollist(content,worksheet3,cell_format)
        print("##### Gathering VMs List #####")
        vmlist(content, worksheet4,cell_format)
        print("##### Gathering Datastores List #####")
        datastoreslist(content,worksheet5, cell_format)
        print("##### Gathering DVSswitch List #####")
        dvslist(content,worksheet6, cell_format)
        print("##### Gathering DVportgroup List #####")
        pglist(content,worksheet7, cell_format)
        workbook.close()
