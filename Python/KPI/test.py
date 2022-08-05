from shareplum import Site
from requests_ntlm import HttpNtlmAuth

auth = HttpNtlmAuth('DIR\\username', 'password')
site = Site('https://abc.com/sites/MySharePointSite/', auth=auth)
sp_list = site.List('list name')
data = sp_list.GetListItems('All Items', row_limit=200)