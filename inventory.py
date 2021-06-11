from pyvim.connect import SmartConnect
from pyVmomi import vim
import xlsxwriter
import getpass



def dclist(content,worksheet):
    viewType = [vim.Datacenter]
    container = content.rootFolder
    recursive = True  # whether we should look into it recursively
    containerView = content.viewManager.CreateContainerView(container, viewType, recursive)
    children = containerView.view
    for child in children:
        print(child)


def vmlist(content, worksheet4):
    worksheet4.write('A1', 'VMName')
    worksheet4.write('B1', 'PowerState')
    worksheet4.write('C1', 'overallStatus')
    worksheet4.write('D1', 'vcpus')
    worksheet4.write('E1', 'numCoresPerSocket')
    worksheet4.write('F1', 'Memory')
    worksheet4.write('G1', 'ESXiHost')
    worksheet4.write('H1', 'VMDKs')
    worksheet4.write('I1', 'vNICs')
    worksheet4.write('J1', 'IP_Address')
    worksheet4.write('K1', 'Hardware_Version')
    worksheet4.write('L1', 'ResourcePool')
    worksheet4.write('M1', 'VMtools_Status')
    worksheet4.write('N1', 'VMtools_version')
    worksheet4.write('O1', 'VMtools_version1state')
    worksheet4.write('P1', 'VMtools_version2state')
    worksheet4.write('Q1', 'guestFullName')
    worksheet4.write('R1', 'HAprotected')
    worksheet4.write('S1', 'snapshots')
    worksheet4.write('T1', 'cpuHotAddEnabled')
    worksheet4.write('U1', 'memoryHotAddEnabled')
    worksheet4.write('V1', 'guestMemoryUsage')
    worksheet4.write('W1', 'overallCpuUsage')
    worksheet4.write('X1', 'maxCpuUsage')
    worksheet4.write('Y1', 'maxMemoryUsage')
    worksheet4.write('Z1', 'cpualloc_limit')
    worksheet4.write('AA1', 'cpuReservation')
    worksheet4.write('AB1', 'cpualloc_share_level')
    worksheet4.write('AC1', 'cpualloc_shares')
    worksheet4.write('AD1', 'memalloc_limit')
    worksheet4.write('AE1', 'memoryReservation')
    worksheet4.write('AF1', 'memalloc_share_level')
    worksheet4.write('AG1', 'memalloc_shares')
    worksheet4.write('AH1', 'balloonedMemory')
    worksheet4.write('AI1', 'consumedOverheadMemory')
    worksheet4.write('AJ1', 'swappedMemory')
    worksheet4.write('AK1', 'swapFile')
    worksheet4.write('AL1', 'guestHeartbeatStatus')
    worksheet4.write('AM1', 'uptimeSeconds')
    worksheet4.write('AN1', 'minRequiredEVCModeKey')
    worksheet4.write('AO1', 'faultToleranceState')
    worksheet4.write('AP1', 'changeTrackingEnabled')
    worksheet4.write('AQ1', 'toolsInstallerMounted')
    worksheet4.write('AR1', 'vFlashCacheAllocation')
    worksheet4.write('AS1', 'installBootRequired')
    worksheet4.write('AT1', 'uuid')
    worksheet4.write('AU1', 'vmPathName')
    worksheet4.write('AV1', 'MksConnections')
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
        worksheet4.write(r, c, configure.config.memorySizeMB)
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
        worksheet4.write(r, c, configure.quickStats.guestMemoryUsage)
        c += 1
        worksheet4.write(r, c, configure.quickStats.overallCpuUsage)
        c += 1
        worksheet4.write(r, c, configure.runtime.maxCpuUsage)
        c += 1
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




if __name__ == '__main__':
    Hostname = input("Enter the vCenter HostName: ")
    vcuser = input("Enter the vCenter UserName: ")
    vcpasswd = getpass.getpass(prompt='Enter vCenter Password: ')
    service_instance = SmartConnect(host=Hostname, user=vcuser, pwd=vcpasswd)
    content = service_instance.RetrieveContent()
    if content:
        workbook = xlsxwriter.Workbook(Hostname + '@' + vcuser + '.xlsx')
        worksheet = workbook.add_worksheet('Datacenters')
        worksheet1 = workbook.add_worksheet('Clusters')
        worksheet2 = workbook.add_worksheet('Host System')
        worksheet3 = workbook.add_worksheet('Resource_Pools')
        worksheet4 = workbook.add_worksheet('Virtual_Machines')
        worksheet5 = workbook.add_worksheet('Datastores')
        worksheet6 = workbook.add_worksheet('vSwitch_DVS')
        worksheet7 = workbook.add_worksheet('Portgroup')
        worksheet8 = workbook.add_worksheet('Folder')
        dclist(content,worksheet)
        #vmlist(content, worksheet4)
        workbook.close()