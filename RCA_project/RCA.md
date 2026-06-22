# Incident Investigation & Root Cause Analysis (RCA)
## SLES for SAP High-Availability Cluster Outage & Reboot Loop

---

### Document Metadata
| Attribute | Details |
| :--- | :--- |
| **Incident Date** | 2026-04-01 |
| **Severity** | P1 - Critical Outage |
| **Incident Window** | 10:06 IST - 11:25 IST (79 minutes) |
| **Impacted Infrastructure** | SUSE Linux Enterprise Server (SLES) 15 SP4 HA Cluster (`hacluster`) |
| **Impacted Nodes** | `iprslesdqap` (Node 1, 192.168.3.31), `iprdqapd` (Node 2, 192.168.3.63) |
| **Lead SRE/Systems Engineer** | Navneet |

---

## 1. Problem Statement
Between 10:06 IST and 11:25 IST on April 1, 2026, the primary high-availability (HA) cluster node `iprslesdqap` (192.168.3.31) entered an unrecoverable reboot loop, undergoing five (5) consecutive SBD (Storage-Based Death) watchdog-triggered reboots. 

During the incident:
1. The secondary node `iprdqapd` (192.168.3.63), acting as the Cluster Designated Coordinator (DC), attempted to fence the primary node at 10:14:47 IST, but the STONITH fencing command failed.
2. The network subsystem on the secondary node `iprdqapd` had its network daemon (`wicked`) restarted at 10:14:19 IST, 28 seconds prior to the fencing failure.
3. Post-restart diagnostic pings from `iprdqapd` showed 100% packet loss to the default gateway (192.168.1.1) and DNS servers, indicating network isolation, while the primary node `iprslesdqap` maintained perfect connectivity.
4. The cluster remained in an unstable split-brain state with resource management disabled until manual intervention placed the cluster in maintenance mode.

---

## 2. Incident Timeline
The table below represents the chronological sequence of events reconstructed from system logs, pacemaker journals, `wtmp` entries, and supportconfig bundles:

| Timestamp (IST) | Node | Event | Source Log / Reference |
| :--- | :--- | :--- | :--- |
| **06:02:27** | `iprdqapd` (Node 2) | Node rebooted and brought back online. | `/var/log/wtmp` (`last -wxF`) |
| **06:02:48** | `iprdqapd` (Node 2) | Runlevel 5 reached, Pacemaker cluster services started. | System logs / pacemaker status |
| **10:06:43** | `iprslesdqap` (Node 1) | **FIRST REBOOT**: SBD watchdog triggers fencing, ending a 146-day system uptime. | `/var/log/wtmp` (`last -wxF`) |
| **10:07:04** | `iprslesdqap` (Node 1) | System boots up, reaches Runlevel 5. | `/var/log/wtmp` |
| **10:14:19** | `iprdqapd` (Node 2) | Network interface daemon `wicked` is manually restarted (PID 50401). | `systemctl status network.service` |
| **10:14:47** | `iprdqapd` (Node 2) | DC attempts to fence `iprslesdqap` (Node 1) via SBD STONITH API. **Fencing fails** due to network isolation. | `crm_mon -r -1` |
| **10:18:21** | `iprslesdqap` (Node 1) | **SECOND REBOOT**: System resets via watchdog (uptime: 11 mins). | `/var/log/wtmp` |
| **10:52:22** | `iprslesdqap` (Node 1) | **THIRD REBOOT**: System resets via watchdog (uptime: 34 mins). | `/var/log/wtmp` |
| **11:05:18** | `iprslesdqap` (Node 1) | **FOURTH REBOOT**: System resets via watchdog (uptime: 12 mins). | `/var/log/wtmp` |
| **11:25:51** | `iprslesdqap` (Node 1) | **FIFTH REBOOT**: System resets via watchdog (uptime: 20 mins). | `/var/log/wtmp` |
| **11:26:12** | `iprslesdqap` (Node 1) | Node 1 successfully completes boot sequence. | `/var/log/wtmp` |
| **11:26:25** | `iprslesdqap` (Node 1) | Pacemaker starts local cluster resources. SAP Application VMs (`p-vm-*`) successfully start. | `systemctl status pacemaker.service` |
| **11:27:24** | `iprslesdqap` (Node 1) | Cluster administration manual override: maintenance-mode enabled (`maintenance-mode=true`). | `cibadmin` configuration log |
| **11:36:31** | Both Nodes | Corosync token timeout of 5000ms reached; ring membership breaks. | `/var/log/messages` (Corosync) |
| **11:36:35** | Both Nodes | Consensus reached after 36000ms. 2-node membership re-established. | `/var/log/messages` (Quorum) |

---

## 3. Technical Findings & Root Cause Analysis

### Finding 1: The Network Isolation of Node 2 (`iprdqapd`)
The root cause sequence began when the `wicked` network service on Node 2 was restarted at 10:14:19 IST. 
* Although local network interfaces came back up, Node 2 suffered routing tables or network bridge state corruption. It lost all connectivity to the wider subnet.
* Diagnostic ping tests from Node 2 to default gateway (`192.168.1.1`), internal DNS (`192.168.1.122`), and public DNS (`8.8.8.8`) returned **100% packet loss**.
* The ARP cache table on Node 2 flagged MAC addresses for default routers and neighboring DNS servers as `STALE` or `FAILED`.
* Conversely, Node 1 (`iprslesdqap`) maintained healthy interface routing and registered 0% packet loss to the default gateway and DNS servers.

### Finding 2: The Fencing Command Failure
Because Node 2 (`iprdqapd`) was the Designated Coordinator (DC) of the Pacemaker cluster, when it lost corosync heartbeat packets from Node 1 due to the network split, it acted to protect data integrity by executing a STONITH fence (reboot) against Node 1 at 10:14:47 IST.
* This fencing action failed because Node 2 was network-isolated and could not communicate with the SBD block storage devices or physical fencing interfaces to issue the fencing command.
* Pacemaker registered the failed fencing action, leaving the cluster state in a dangerous partition.

### Finding 3: SBD Watchdog Self-Fencing & The Reboot Loop
Node 1 (`iprslesdqap`) was running an active hardware watchdog integration (`/dev/watchdog`) linked to the Storage-Based Death (SBD) daemon.
* When the corosync token timeout (5000ms) expired, Node 1 recognized that it was isolated from Node 2.
* Under SLES HA architecture, if a node loses cluster communication (and quorum), it must verify its own state via the SBD disk heartbeats. Because Node 2 had restarted network interfaces and disrupted paths, and the cluster suffered split-brain symptoms, Node 1's local SBD daemon failed to update its watchdog timer in time.
* The hardware watchdog triggered a hard system reboot to prevent split-brain data corruption.
* Once Node 1 rebooted and came back up, it attempted to rejoin the cluster. However, because Node 2 remained network-isolated with corrupted routing, Node 1 again failed to establish stable cluster communications. The corosync token expired again, triggering subsequent SBD watchdog reboots (a total of 5 times in 79 minutes).

### Finding 4: Aggressive Timeout Settings & Configuration Mismatches
Two configuration vulnerabilities sustained this reboot loop:
1. **Aggressively Low Corosync Token Timeout**: The corosync configuration had token timeout set to `5000ms` (5 seconds). On virtualized hardware or during minor network blips, 5 seconds is too short, causing corosync to false-alarm on cluster communication loss.
2. **SBD Watchdog Timeout Mismatch**: The SBD config file `/etc/sysconfig/sbd` declared `SBD_WATCHDOG_TIMEOUT=5`, whereas the actual SBD shared disk header had a watchdog timeout configured for `60s`. This discrepancy resulted in Pacemaker and the kernel SBD driver having mismatched expectations of fencing timelines, causing premature watchdog resets before the node could properly negotiate state transitions or finish network initialization.

---

## 4. System Impact
- **Node Down-time**: Node 1 (`iprslesdqap`) was completely unavailable for cluster workloads for 79 minutes.
- **Resource Consolidation**: All cluster resources (SAP application virtual machines: `p-vm-iprappprd`, `p-vm-iprtedd`, `p-vm-iprsolman`, `p-vm-iprweb`) were consolidated on Node 1 after it finally stabilized. Node 2 (`iprdqapd`) runs zero VMs due to its degraded network state.
- **Service Degradation**: Four (4) critical VMs (`iprpootpd`, `iprappdev`, `iprappqua`, `iprpootdv`) were stopped during the reboots and remained offline.
- **Operational Status**: The cluster was placed in manual `maintenance-mode`, disabling automatic resource control and failover capabilities.

---

## 5. Action Plan & Recommendations

### Immediate Remediation Tasks
1. **Resolve Node 2 Network Bridge / Routing**:
   - Inspect the physical network interfaces, switch ports, and VLAN configurations for bond `bond0`/`bond1` and bridge `br0` on Node 2.
   - Verify Node 2 can ping the default gateway (`192.168.1.1`) and DNS server (`192.168.1.122`) with 0% packet loss.
2. **Clear Cluster Fencing History**:
   - Reset the failed fencing history recorded in Pacemaker to allow normal transition logic:
     ```bash
     stonith_admin --cleanup --history
     ```
3. **Clean Up Location Constraints**:
   - Delete manual CLI-imposed location rules that prevent VMs from failback:
     ```bash
     crm configure delete cli-ban-p-vm-iprappprd-on-iprdqapd
     crm configure delete cli-prefer-p-vm-iprappprd
     ```
4. **Remove Cluster Maintenance Mode**:
   - Re-enable Pacemaker control:
     ```bash
     crm configure property maintenance-mode=false
     ```

### High-Availability Configuration Hardening
5. **Increase Corosync Token Timeout**:
   - Edit `/etc/corosync/corosync.conf` on both nodes and increase the `token` timeout value from `5000` to `10000` (10 seconds) to tolerate transient hypervisor or network latencies.
6. **Align SBD Watchdog Settings**:
   - Ensure the software SBD configurations match the physical disk header settings. Update `/etc/sysconfig/sbd` on both nodes:
     ```env
     SBD_WATCHDOG_TIMEOUT=60
     ```
7. **Configure SBD Start Delay**:
   - Enable `SBD_DELAY_START=yes` in `/etc/sysconfig/sbd`. This ensures that on boot, SBD waits for network interfaces and corosync to fully initialize before arming the watchdog, avoiding boot loops during network start-up delays.

### Long-Term Architectural Improvements
8. **Implement QDevice (Quorum Device)**:
   - For a 2-node cluster, a single node loss represents 50% quorum loss. Adding a lightweight Corosync QDevice on a third network segment provides a tie-breaker vote, preventing split-brain conditions and unwanted fencing.
9. **Dual-Ring Network Path Redundancy**:
   - Configure a secondary Corosync ring (Ring 1) over a separate physical network interface or subnet to ensure heartbeat traffic continues even if `br0`/`bond1` experiences an outage.

---

## 6. Diagnostic Evidence & Log Snippets

### A. Corosync Token Timeout Logs (Simultaneous on both nodes)
*Log excerpt showing the membership loss on Node 1 and Node 2 simultaneously at 11:36:31:*

**Node 1 (`iprslesdqap`):**
```text
2026-04-01T11:36:31.689105+05:30 iprslesdqap corosync[3371]:   [TOTEM ] A processor failed, forming new configuration: token timed out (5000ms), waiting 36000ms for consensus.
2026-04-01T11:36:35.182811+05:30 iprslesdqap corosync[3371]:   [TOTEM ] A new membership (192.168.3.31:2094) was formed. Members
2026-04-01T11:36:35.184134+05:30 iprslesdqap corosync[3371]:   [QUORUM] Members[2]: 1 2
2026-04-01T11:36:35.184251+05:30 iprslesdqap corosync[3371]:   [MAIN  ] Completed service synchronization, ready to provide service.
```

**Node 2 (`iprdqapd`):**
```text
2026-04-01T11:36:31.689374+05:30 iprdqapd corosync[3964]:   [TOTEM ] A processor failed, forming new configuration: token timed out (5000ms), waiting 36000ms for consensus.
2026-04-01T11:36:35.182942+05:30 iprdqapd corosync[3964]:   [TOTEM ] A new membership (192.168.3.31:2094) was formed. Members
2026-04-01T11:36:35.184828+05:30 iprdqapd corosync[3964]:   [QUORUM] Members[2]: 1 2
2026-04-01T11:36:35.184910+05:30 iprdqapd corosync[3964]:   [MAIN  ] Completed service synchronization, ready to provide service.
```

---

### B. Failed Fencing (STONITH API Error)
*Execution output of `crm_mon` illustrating the failed fencing attempt initiated by Node 2:*
```text
Failed Fencing Actions:
  * reboot of 1 failed: delegate=, client=stonith-api.51209, origin=iprdqapd, last-failed='2026-04-01 10:14:47 +05:30'
```

---

### C. System Boot History on Node 1 (`iprslesdqap`)
*Wtmp records demonstrating the five consecutive reboots between 10:06 and 11:25:*
```text
runlevel (to lvl 5)   5.14.21-150400.24.179-default Wed Apr  1 11:26:12 2026   still running
reboot   system boot  5.14.21-150400.24.179-default Wed Apr  1 11:25:51 2026   still running
runlevel (to lvl 5)   5.14.21-150400.24.179-default Wed Apr  1 11:05:38 2026 - Wed Apr  1 11:26:12 2026  (00:20)
reboot   system boot  5.14.21-150400.24.179-default Wed Apr  1 11:05:18 2026   still running
runlevel (to lvl 5)   5.14.21-150400.24.179-default Wed Apr  1 10:52:43 2026 - Wed Apr  1 11:05:38 2026  (00:12)
reboot   system boot  5.14.21-150400.24.179-default Wed Apr  1 10:52:22 2026   still running
runlevel (to lvl 5)   5.14.21-150400.24.179-default Wed Apr  1 10:18:42 2026 - Wed Apr  1 10:52:43 2026  (00:34)
reboot   system boot  5.14.21-150400.24.179-default Wed Apr  1 10:18:21 2026   still running
runlevel (to lvl 5)   5.14.21-150400.24.179-default Wed Apr  1 10:07:04 2026 - Wed Apr  1 10:18:42 2026  (00:11)
reboot   system boot  5.14.21-150400.24.179-default Wed Apr  1 10:06:43 2026   still running
runlevel (to lvl 5)   5.14.21-150400.24.179-default Thu Nov  6 07:07:32 2025 - Wed Apr  1 10:07:04 2026 (146+02:59)
```

---

### D. Node 2 (`iprdqapd`) Network Diagnostic Logs
*Proof of the wicked restart at 10:14:19:*
```text
wicked.service - wicked managed network interfaces
    Loaded: loaded (/usr/lib/systemd/system/wicked.service; enabled; vendor preset: disabled)
    Active: active (exited) since Wed 2026-04-01 10:14:19 IST; 1h 21min ago
   Process: 50401 ExecStart=/usr/sbin/wicked --systemd ifup all (code=exited, status=0/SUCCESS)
```

*Proof of 100% packet loss to the default gateway and internal DNS:*
```bash
# Gateway Connection Test
iprdqapd:~ # ping -n -c1 -W1 192.168.1.1
PING 192.168.1.1 (192.168.1.1) 56(84) bytes of data.
--- 192.168.1.1 ping statistics ---
1 packets transmitted, 0 received, 100% packet loss, time 0ms

# DNS Connection Test
iprdqapd:~ # ping -n -c1 -W1 192.168.1.122
PING 192.168.1.122 (192.168.1.122) 56(84) bytes of data.
--- 192.168.1.122 ping statistics ---
1 packets transmitted, 0 received, 100% packet loss, time 0ms
```

*ARP state showing stale routers and neighbor validation failures:*
```text
iprdqapd:~ # ip -stats neighbor
192.168.1.122 dev bond1 lladdr 00:50:56:a9:1c:79 used 131/126/93 probes 1 STALE
192.168.1.1 dev br0 lladdr d4:76:a0:56:73:0e ref 1 used 39/26/39 probes 3 REACHABLE
fe80::6a4f:64ff:fe8e:1e13 dev br0  router used 174/241/173 probes 3 FAILED
fe80::6a4f:64ff:fe8e:4413 dev br0  router used 192/259/191 probes 3 FAILED
```
