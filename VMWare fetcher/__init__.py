import json
import ssl

from pyVim import connect
from pyVmomi import vmodl
from pyVmomi import vim


def main():
    """
   Simple command-line program for listing all ESXi datastores and their
   associated devices
   """
    host = 'mow03-vvs01.cherkizovsky.net'
    user = 'CHERKIZOVSKY\\TECH_ZAB_VM'
    pwd = '7Ns4vn44B447'
    sslContext = None
    sslContext = ssl.SSLContext(ssl.PROTOCOL_SSLv23)
    sslContext.verify_mode = ssl.CERT_NONE
    

    try:
        service_instance = connect.SmartConnect(host,
                                                user,
                                                pwd,
                                                port=433,
                                                sslContext=sslContext)
        if not service_instance:
            print("Could not connect to the specified host using specified "
                  "username and password")
            return -1

        atexit.register(connect.Disconnect, service_instance)

        content = service_instance.RetrieveContent()
        # Search for all ESXi hosts
        objview = content.viewManager.CreateContainerView(content.rootFolder,
                                                          [vim.HostSystem],
                                                          True)
        esxi_hosts = objview.view
        objview.Destroy()

        datastores = {}
        for esxi_host in esxi_hosts:
            print("{}\t{}\t\n".format("ESXi Host:    ", esxi_host.name))

            # All Filesystems on ESXi host
            storage_system = esxi_host.configManager.storageSystem
            host_file_sys_vol_mount_info = \
                storage_system.fileSystemVolumeInfo.mountInfo

            datastore_dict = {}
            # Map all filesystems
            for host_mount_info in host_file_sys_vol_mount_info:
                # Extract only VMFS volumes
                if host_mount_info.volume.type == "VMFS":

                    extents = host_mount_info.volume.extent
                    if not args.json:
                        print_fs(host_mount_info)
                    else:
                        datastore_details = {
                            'uuid': host_mount_info.volume.uuid,
                            'capacity': host_mount_info.volume.capacity,
                            'vmfs_version': host_mount_info.volume.version,
                            'local': host_mount_info.volume.local,
                            'ssd': host_mount_info.volume.ssd
                        }

                    extent_arr = []
                    extent_count = 0
                    for extent in extents:
                        if not args.json:
                            print("{}\t{}\t".format(
                                "Extent[" + str(extent_count) + "]:",
                                extent.diskName))
                            extent_count += 1
                        else:
                            # create an array of the devices backing the given
                            # datastore
                            extent_arr.append(extent.diskName)
                            # add the extent array to the datastore info
                            datastore_details['extents'] = extent_arr
                            # associate datastore details with datastore name
                            datastore_dict[host_mount_info.volume.name] = \
                                datastore_details
                    if not args.json:
                        print

            # associate ESXi host with the datastore it sees
            datastores[esxi_host.name] = datastore_dict

        if args.json:
            print(json.dumps(datastores))

    except vmodl.MethodFault as error:
        print("Caught vmodl fault : " + error.msg)
        return -1

    return 0


# Start program
if __name__ == "__main__":
    main()