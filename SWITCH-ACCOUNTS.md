# How to Switch NetSuite Accounts

## 📝 Simple Config Switching

Your backend is configured to work with one account at a time. To switch accounts, edit `backend/netsuite_config.json`.

---

## 🏢 Production Account (589861)

Replace `backend/netsuite_config.json` with:

```json
{
    "account_id": "589861",
    "consumer_key": "a432d93c007d27337151ee010d734bb9858556addc3d5961923fbf15ef2d8246",
    "consumer_secret": "953405056dc879569d03e074422ca1792bf5c34f4d724fdafd001dbf6a8e5df3",
    "token_id": "fd15642ac1360a727dee4076a137e1825a987e4b07d6216a1f7674311dfd7db0",
    "token_secret": "8138f3291f9a1fa97dd8aacda7bca1a0d3bff876dc3e732a35abbf8366608d2d"
}
```

---

## 🧪 Sandbox Account (TSTDRV2320150)

Replace `backend/netsuite_config.json` with:

```json
{
    "account_id": "TSTDRV2320150",
    "consumer_key": "2ac58543b20fe325f0735054bb09a08a732415337e2909b0798b791323655fec",
    "consumer_secret": "e967c377baf509e69ae9c11fd29586f5e04e9e8708d984db037c0f187f8acdd5",
    "token_id": "356e6b89fc14256cd2a5842cdef0ab4486d9fcd2f37e6c5f6e2cdadb5930e66b",
    "token_secret": "a57c99238b8f7240dae3989e99d0f998ad89ab983a72bd42a6a846d9ae51fd04"
}
```

---

## 🔄 After Editing Config:

```bash
cd "/Users/chriscorcoran/Documents/Cursor/NetSuite Formulas Revised"
./restart-servers.sh
```

Wait 15 seconds for servers to start, then test in Excel!

---

## 🧪 Quick Test:

```bash
# Test connection
curl http://localhost:5002/test

# Test formula
curl "http://localhost:5002/account/4000/name"
```

---

**Last Updated:** Nov 30, 2025  
**Current Account:** Production (589861)

