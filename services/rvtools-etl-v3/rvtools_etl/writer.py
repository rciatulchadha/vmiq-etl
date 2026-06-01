"""
Snapshot writer v3 — estate-aware.
Upserts estate and vCenter records before writing snapshot data.
"""

import logging
import uuid
from datetime import datetime, timezone

import asyncpg

log = logging.getLogger("rvtools-etl.writer")


class SnapshotWriter:
    def __init__(self, pool: asyncpg.Pool):
        self.pool = pool

    async def write(self, data: dict) -> list[str]:
        snapshot_ids = []
        estate_name  = data.get("estate", "DEFAULT")
        vcenters     = data["vcenters"]

        if not vcenters:
            vcenters = [{"name": "Unknown", "fqdn": "unknown",
                         "estate": estate_name}]

        for vc in vcenters:
            vc_name = vc["name"]
            log.info(f"Writing snapshot — Estate: {estate_name}, vCenter: {vc_name}")

            # Filter data for this vCenter
            # If only one vCenter in file, assign everything to it
            single = len(vcenters) == 1

            vc_clusters   = [c for c in data["clusters"]
                             if single or c.get("vcenter") == vc_name]
            vc_hosts      = [h for h in data["hosts"]
                             if single or h.get("vcenter") == vc_name]
            vc_vms        = [v for v in data["vms"]
                             if single or v.get("vcenter") == vc_name]
            vc_datastores = [d for d in data["datastores"]
                             if single or d.get("vcenter") == vc_name]

            snapshot_id = await self._write_vcenter_snapshot(
                estate_name, vc_name, vc["fqdn"],
                vc_clusters, vc_hosts, vc_vms, vc_datastores
            )
            snapshot_ids.append(snapshot_id)
            log.info(
                f"  Snapshot {snapshot_id[:8]}... — "
                f"{len(vc_vms)} VMs, {len(vc_hosts)} hosts, "
                f"{len(vc_clusters)} clusters, "
                f"{len(vc_datastores)} datastores"
            )

        return snapshot_ids

    async def _write_vcenter_snapshot(
        self,
        estate_name: str,
        vc_name: str,
        vc_fqdn: str,
        clusters: list,
        hosts: list,
        vms: list,
        datastores: list,
    ) -> str:
        async with self.pool.acquire() as conn:
            async with conn.transaction():

                # 1. Upsert estate
                estate_id = await self._upsert_estate(conn, estate_name)

                # 2. Upsert vCenter
                vc_id = await self._upsert_vcenter(
                    conn, vc_name, vc_fqdn, estate_id)

                # 3. Create snapshot
                snapshot_id = str(uuid.uuid4())
                await conn.execute("""
                    INSERT INTO vcenter_snapshots
                        (id, vcenter_id, collected_at, status,
                         vm_count, host_count, cluster_count,
                         source_file)
                    VALUES ($1, $2, NOW(), 'processing',
                            $3, $4, $5, 'rvtools_excel')
                """, snapshot_id, vc_id,
                    len(vms), len(hosts), len(clusters))

                # 4. Write child records
                cluster_id_map = await self._write_clusters(
                    conn, snapshot_id, vc_id, clusters)
                host_id_map = await self._write_hosts(
                    conn, snapshot_id, cluster_id_map, hosts)
                await self._write_vms(
                    conn, snapshot_id, cluster_id_map, host_id_map, vms)
                await self._write_datastores(
                    conn, snapshot_id, datastores)

                # 5. Mark complete
                await conn.execute("""
                    UPDATE vcenter_snapshots
                    SET status = 'complete'
                    WHERE id = $1
                """, snapshot_id)

        return snapshot_id

    async def _upsert_estate(self, conn, name: str) -> str:
        """Insert estate if not exists, return UUID."""
        row = await conn.fetchrow(
            "SELECT id FROM estates WHERE name = $1", name)
        if row:
            return str(row["id"])
        estate_id = str(uuid.uuid4())
        await conn.execute("""
            INSERT INTO estates (id, name, created_at)
            VALUES ($1, $2, NOW())
        """, estate_id, name)
        log.info(f"  Created estate: {name}")
        return estate_id

    async def _upsert_vcenter(self, conn, name: str,
                               fqdn: str, estate_id: str) -> str:
        """Insert vCenter if not exists, return UUID."""
        row = await conn.fetchrow(
            "SELECT id FROM vcenters WHERE name = $1", name)
        if row:
            # Update estate_id if missing
            await conn.execute("""
                UPDATE vcenters SET estate_id = $2
                WHERE id = $1 AND estate_id IS NULL
            """, str(row["id"]), estate_id)
            return str(row["id"])
        vc_id = str(uuid.uuid4())
        await conn.execute("""
            INSERT INTO vcenters
                (id, name, fqdn, estate_id, created_at)
            VALUES ($1, $2, $3, $4, NOW())
        """, vc_id, name, fqdn, estate_id)
        log.info(f"  Created vCenter: {name}")
        return vc_id

    async def _write_clusters(self, conn, snapshot_id: str,
                               vc_id: str, clusters: list) -> dict:
        cluster_id_map = {}
        records = []
        for c in clusters:
            cid = str(uuid.uuid4())
            cluster_id_map[c["name"]] = cid
            records.append((
                cid, snapshot_id, vc_id,
                c["name"],
                c.get("cpu_total_mhz") or 0,
                c.get("mem_total_mb") or 0,
                c.get("host_count") or 0,
                c.get("vm_count") or 0,
            ))
        if records:
            await conn.copy_records_to_table(
                "clusters", records=records,
                columns=["id", "snapshot_id", "vcenter_id", "name",
                         "cpu_total_mhz", "mem_total_mb",
                         "host_count", "vm_count"])
        return cluster_id_map

    async def _write_hosts(self, conn, snapshot_id: str,
                            cluster_id_map: dict,
                            hosts: list) -> dict:
        host_id_map = {}
        records = []
        for h in hosts:
            hid = str(uuid.uuid4())
            host_id_map[h["name"]] = hid
            cluster_id = cluster_id_map.get(h.get("cluster"))
            records.append((
                hid, snapshot_id, cluster_id,
                h["name"],
                h.get("esxi_version"),
                h.get("esxi_build"),
                h.get("model"),
                h.get("vendor"),
                h.get("cpu_sockets") or 0,
                h.get("cpu_cores") or 0,
                h.get("cpu_threads") or 0,
                h.get("mem_total_mb") or 0,
                h.get("powerstate", "unknown"),
                h.get("connection_state", "unknown"),
                h.get("is_in_maintenance", False),
                h.get("is_intel_cpu", True),
                h.get("mgmt_ip"),
            ))
        if records:
            await conn.copy_records_to_table(
                "esx_hosts", records=records,
                columns=["id", "snapshot_id", "cluster_id", "name",
                         "esxi_version", "esxi_build", "model", "vendor",
                         "cpu_sockets", "cpu_cores", "cpu_threads",
                         "mem_total_mb", "powerstate", "connection_state",
                         "is_in_maintenance", "is_intel_cpu", "mgmt_ip"])
        return host_id_map

    async def _write_vms(self, conn, snapshot_id: str,
                          cluster_id_map: dict,
                          host_id_map: dict,
                          vms: list) -> None:
        BATCH = 500
        now = datetime.now(timezone.utc)
        columns = [
            "id", "snapshot_id", "host_id", "cluster_id",
            "name", "powerstate", "os_type", "os_fullname",
            "vcpus", "mem_mb", "disk_total_gb", "ip_address",
            "tools_status", "tools_version", "hw_version",
            "has_snapshots", "has_rdm", "has_usb", "has_cdrom",
            "is_suspended", "net_adapter", "created_date", "collected_at",
        ]

        def make_record(v):
            return (
                str(uuid.uuid4()), snapshot_id,
                host_id_map.get(v.get("host")),
                cluster_id_map.get(v.get("cluster")),
                v["name"], v.get("powerstate", "unknown"),
                v.get("os_type"), v.get("os_fullname"),
                v.get("vcpus") or 0,
                v.get("mem_mb") or 0,
                v.get("disk_total_gb") or 0,
                v.get("ip_address"),
                v.get("tools_status"), v.get("tools_version"),
                v.get("hw_version"),
                v.get("has_snapshots", False),
                v.get("has_rdm", False),
                v.get("has_usb", False),
                v.get("has_cdrom", False),
                v.get("is_suspended", False),
                v.get("net_adapter"),
                v.get("created_date"),
                now,
            )

        for i in range(0, len(vms), BATCH):
            batch = vms[i:i + BATCH]
            await conn.copy_records_to_table(
                "virtual_machines",
                records=[make_record(v) for v in batch],
                columns=columns)
            log.debug(
                f"  VMs written: {min(i+BATCH, len(vms))}/{len(vms)}")

    async def _write_datastores(self, conn, snapshot_id: str,
                                  datastores: list) -> None:
        records = [(
            str(uuid.uuid4()), snapshot_id,
            d["name"], d.get("type"),
            d.get("capacity_gb") or 0,
            d.get("free_gb") or 0,
            d.get("used_pct") or 0,
            d.get("vm_count") or 0,
        ) for d in datastores]
        if records:
            await conn.copy_records_to_table(
                "datastores", records=records,
                columns=["id", "snapshot_id", "name", "type",
                         "capacity_gb", "free_gb",
                         "used_pct", "vm_count"])
