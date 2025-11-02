# Episode Datasheets

This folder contains CSV episode data used by Episode Organiser.

- `thomas_&_friends_(1984).csv`: episode data based on https://en.wikipedia.org/wiki/List_of_Thomas_%26_Friends_episodes, movie entries are based on https://www.imdb.com/list/ls098617186/.

## Expected CSV Format

- Header row is required: `ep_no,series_ep_code,title,air_date`
- Rows are comma‑separated, typically quoted (RFC‑4180 style):

```
"001","s01e01","Thomas & Gordon","1984-10-09"
"321.5","s00e04","Hero of the Rails","2009-09-08"
```

- Column meanings:
- `ep_no`: overall episode number. Use decimals for films/specials if needed (e.g. `321.5`).
- `series_ep_code`: `sXXeXX` for episodes, `s00eXX` for films/specials.
- `title`: episode or film title.
- `air_date`: ISO date `YYYY-MM-DD`. Leave blank if unknown.

- Encoding: UTF‑8 is recommended.
- Extra columns are ignored.
- Season mapping: episodes go to `Season N` from `sXX`, films/specials go to `Season 0`.
