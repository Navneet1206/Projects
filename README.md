# Systems Engineering & DevOps Portfolio
## Navneet | Systems Engineer, DevOps & SRE Specialist

Welcome to my systems engineering and DevOps portfolio repository. This repository showcases real-world implementations, automation tools, and diagnostics reports demonstrating my capabilities in infrastructure automation, system security hardening, and high-availability cluster incident response.

---

## Featured Projects

### 1. [Linux Server Hardening Automation (Ansible)](Navneet-Server-Hardning)
Automated playbooks designed to harden Linux operating systems (Ubuntu 22.04 LTS and openSUSE Leap) in alignment with industry-standard CIS/NIST/STIG security baselines. 

*   **Key Features**:
    *   **Access Hardening**: Disables remote root SSH access, enforces public-key authentication, and configures password strength policies via PAM/`pwquality`.
    *   **Firewall Enforcement**: Implements network ingress security using UFW on Debian-based systems and firewalld on SUSE systems.
    *   **Insecure Daemon Clean-up**: Disables and masks legacy protocols (Telnet, FTP, Rsh, TFTP, xinetd).
    *   **System Auditing**: Configures persistent logging through `journald`, log retention policies in `logrotate`, and audit rules via `auditd`.
    *   **Verification**: Automatically triggers security compliance analysis using the Lynis auditing framework.
*   **Technologies Used**: Ansible Core, Python, Bash, UFW, Firewalld, PAM, Auditd, Lynis.
*   **Links**:
    *   [Project Directory](Navneet-Server-Hardning)
    *   [Setup & Installation Guide](Navneet-Server-Hardning/docs/setup_guide.md)
    *   [Testing & Results Summary](Navneet-Server-Hardning/docs/testing_and_results.md)
    *   [Architecture Overview](Navneet-Server-Hardning/docs/project_overview.md)

---

### 2. [SLES HA Cluster Incident Analysis & Root Cause Analysis (RCA)](RCA_project)
A comprehensive post-mortem investigation and Root Cause Analysis (RCA) report analyzing a critical P1 outage on a 2-node High-Availability (HA) cluster running SUSE Linux Enterprise Server (SLES) for SAP Applications.

*   **Key Features**:
    *   **Reconstructed Event Timeline**: Correlated systems logs, wtmp boots history, pacemaker journals, and network status files to construct a millisecond-accurate timeline.
    *   **Root Cause Diagnosis**: Diagnosed local bridge and routing failures triggered by a service restart, leading to network isolation, STONITH fencing failures, and a 5x watchdog-triggered reboot loop.
    *   **Configuration Alignment**: Identified critical cluster timing vulnerabilities, including aggressive corosync token timeouts and SBD watchdog config mismatches against physical block device headers.
    *   **Remediation & Action Plan**: Formulated immediate cluster recovery procedures and long-term architectural upgrades (e.g., Corosync QDevice and dual-ring networking).
*   **Technologies Used**: SLES for SAP, Pacemaker, Corosync, SBD (Storage-Based Death) fencing, STONITH, `wicked` network service, Supportconfig logs.
*   **Links**:
    *   [Project Directory](RCA_project)
    *   [Full Incident RCA Report](RCA_project/RCA.md)

---

## Technical Skills Profile

*   **Operating Systems**: SUSE Linux Enterprise Server (SLES) 12/15, openSUSE Leap, Ubuntu Server, Red Hat Enterprise Linux (RHEL).
*   **Infrastructure as Code (IaC)**: Ansible playbooks, roles, variable overrides, handlers, and inventory management.
*   **High-Availability Clustering**: Pacemaker, Corosync, SBD disk/diskless fencing, STONITH, QDevice quorum nodes.
*   **Linux Networking & Core Services**: Network bridging, bonding, `wicked`, `systemd-networkd`, SSH, PAM, DNS resolution, ARP validation.
*   **Security & Auditing**: CIS Benchmarks, auditd rules, log rotation, Lynis security scanning.
*   **Troubleshooting & SRE**: Supportconfig archive analysis, post-mortem incident documentation, journald log investigation.
