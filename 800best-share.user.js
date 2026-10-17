// ==UserScript==
// @name         800BEST - Thống kê tổng quan & Shipper (bản chia sẻ)
// @namespace    local.800best.report.share
// @version      3.15.1
// @description  Dashboard hub: tỷ lệ quét giao, POD2400, backlog, tách TikTok Shop, tiến độ từng shipper; lưu lịch sử theo ngày để xem tuần/tháng; tab ANM tự lưu bản chụp theo mốc giờ cố định (10:30, 13:30, 15:30, 18:30, 23:30) để so sánh đúng giờ với hôm qua/tuần trước; xuất Excel. v3.9: bảng xếp hạng thi đua có vương miện (POD2400, POD TikTok, Giao thành công, Tỷ lệ ODR), Backlog 5 ngày, giờ đơn giao thành công đầu tiên, phân tích nguyên nhân tăng/giảm POD (ship, sản lượng, thời tiết), bỏ tab Lịch sử. v3.11: khung COD ship nộp tiền thay biểu đồ 7 ngày (theo khung giờ, cảnh báo ship nộp 0đ hoặc dưới 50%, sao chép gửi nhóm). v3.10: tiêu đề giữa, bảng xếp hạng gọn thêm xu hướng 7 ngày / cần thêm để đạt mục tiêu / chưa quét giao / DM làm, dự báo POD cuối ngày bằng AI (công tắc dưới bảng, nhập khóa API ở Cài đặt).
// @match        https://v9.800best.com/sea-web/*
// @run-at       document-start
// @require      https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js
// @updateURL    https://raw.githubusercontent.com/Erick1939/tool-best/main/800best-share.user.js
// @downloadURL  https://raw.githubusercontent.com/Erick1939/tool-best/main/800best-share.user.js
// @grant        none
// ==/UserScript==

(() => {
  'use strict';

  /* ------------------------------------------------------------------ *
   * State
   * ------------------------------------------------------------------ */
  const REGION = '';
  const LOGO = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAPAAAACKCAIAAAAbjlBiAABPaUlEQVR42u1dd3wVVfY/5947M6+/l94gQARFiqKi2LBg74t9XRd717X3rlhRsXf92daytl11UbBiQbGiItKkh/Tkvbw+M/ee3x/zkryEJJAQVNicTz4f/ZCXeTP3njn3e76nIRFBv/TLpiKs61/1K/qmIbRJP1rHpxMdP0IEgH32fYjYyeJiXz8P69Pb7oMb6/jc63k5oszatR2nCEAACNj6j5i1u873r/9dEAAR4Z9pbddc6nYaJdb4BK7xP70WdLaCiIgAEQgIERiyPrt+5lvW2NA/UPr0NpRSAMAYyzxtOw3FVu3tYstIKSIAznB9Vxj/nIbeeWPb2TL8/TG0UtIxq4jQt5bsz3cg9vLpiChjGFkGEzY2Na6urlu8rDZS17B8xaoFK6ti8VQslUwkEnmhXMHR6zLGbFERCAYKSvLLBxYPGlhSkJvT7mqIm+RqO4/WUaEVAUN49l/vfPTZtz63oVQfaLkiCnp0lzeYn+MfWBTUdKO4pLC4pHBAcVHrZ6RUyJD1dqGVUoyxGR99+fw//x0syJW2/UcfjoScxxqjx//1kH0n7OjcXk9fAiUV55m/+mX+ou9/WvjxZz8snr94wYpVtdVRSCsACV4DBAdqOREQIW1DMg3IwCNyC4LDigoHDR205+5jxm41fLttRmHLcq27Wjs3/+Gn3zz7z//4gl5Sfy4TjQi2VG6Xdv4Zx1UMGrgG5CACxK9nz3nu4degMAi26qvzlyNK5JpgnHNDaLn5weED8saMGX3EgbsOHznM63K3rl2vHZ4ff1z4wtMvQdkQSKXhjzVCRKBrULVqzFbD952wY0+cNmzBysQ5S5rmrM+/efS5/3zyxY/hZMpOpYEJcLlEcS44+CFzymUvNcuAEqUa0/bseUtmz130r/9+EvR7x40cdOSh+048fJ/8nGDWaiuAtYPjX+Ytef7xV6G4ACz55wLSDCFp+kK+Scce6lhPxyy2w9Aet0vkBHlurrKtPvPaiDiRTWARpZQdqapdumL1uzO+v+32h8Zsu/Uxxxx00nEHFuXmOFCEMd6Lb/G4DeHK0XIDdsr8wxWa6ZqMxFwuvad/KjOGGd/74PObbnvyyy/ngNsNfjdzGZov89oraYPdlQOiHLcRELnGWL4fAaRUEdue8dUvM96fffOUpy8659iTTjo86PMppdYR77kNXXhCPDeo/mQKjQxl2g76vby9KRQdThlbSpBSSdm3bhJDAERgHDUOLgOCQJA/Z8GKOZfd/fpL7/3jvCP/fuxhjHEihdhjU60U2bZkUkkp/3CFJslsW/YEsyEASCk556tWV1176yMvvvaxaUleXoYkQSkClLbsxGnv/EoZS+/8CSJwwSHkw9zAyubEhZfe++bbM6+4dNIB++y+JvrsCje2qcSfSqEJpZRSyg5OIPs9dhmAKOOaKEVKKWUpstIi16eXF367dMWkM285+m8XVVbVITLHr/9fElJKcc4/n/Xd3oee88w/p1s+lygIkmUqKRURkVoToiACMmRt0oqMqQMCIiKSSpkW9xliUNGncxYccvw1N932UMpMIYKiTW21RQ8ZHOgb/gYBAJWtFIDwGxD0vvrax6sbw/98avKg0tJ1sRwbOfvRtp5SKc74W9M+OOGsyeE06SX50rKVbXcwhw7byYRARACyTUuaEqQEh4DnjGlC13VCRqSUtEkqyibbEEkqAtJy/dKU19/0eE1d4123Xup2GVmnIv2J+ebudbLnCo0tPgeA6lqn25R+DdK0JTBA2UcEZdwbaesVZV/MnnfCSde8/OLdxXmhXuPpTpxh7MNdWhsfSwAM1521kUpyxt//eNbxZ0yOIhdBzU53dAOQISLaEiBuQrgOzBgoAaGgj2t+j5sLJMVTyWRjNJKKRoFM0L0QzAWvJoQAJdufeCRtGwUT5SUPP/PflLKfuPs6wE72KvuBkREyRMYAqRfvLhJ0QQ0jAgLrpX1Ehsg4rkGxi3X8a2laYJnrEDZo+YAEkBLsFlCrCRAMhGCGxgXKNTwM2zKNkvyZH/1w5rmTX/vnnWzdEN7aKDQkywbT+h2NLzHNplTSWosTgg7058gWL1l22nm3Ri2lh9x2xjBTG0fEuJVMQyyVE9KHVBRvP3q3YUOKPJ5QyeCS4uLcnJwcTdOUtOOx5lVVjSuWrkrHmuYvafxizrxlVTWxqmYIuoXHUHY7G0RKAQIvyHn6ybfLS0qvv/wMpWjNCEVmX5RNybgdS1KvnULG0GXgmmc7grQlJNK9VmhK2klGLe5K5ubFWncINSGrGy85/287bLulZVrrqGSRhElKxSNNRKqpMVLXFJmzeGXVb/WVdU2KkSgIUjs4TwBgpdPaZsX/eefTp15464xJEx3StFdnEAIQMkbRxMT9xx0/cf9kKtVLWrA3RwJaprX1mJHQZegZW/9jKbrhrqeXr6zTyvJs0862bMg5ElhVjYOGlh13yviD958wYovBoVCoq+/delTb/0dj0Z/nLXnnvx898+bHVStreUEugKQsU00EqBQrzL9t6gs7bLPFAfvuQUoi49k67dz8+F23f+7FKZouehGAIwIh+NLlKyc/8EoM2hlTxpndlNhudMVFZx+DvcWwDoc2sLQQoO0YFuuwQwyS4T3GjTjowH3WZ6ebIpGmpvDHn3z93IszPv3sOyjK4YYgqbLtOimFQe+DD/3r8IP3KMjJ6T0/DYCIFEuNGbH54Yfu/Udh5W5eSOfR3n//8zffnMkLQ7L9MYIMpS2hsvb4UyZec+nJWwwpbyNRpCQCxHYREiIgIFKEAJxzv8+/8w5b77zD1pOOP3Tq1Oce/7/XsLiEaby9Tiuui3SUrr/j6T1238mla0oplqV0ztVHDx8yeviQ9VmF+QsW3/XAy7Hsw50AGUBSDsrLOe6IA/vKjvSE5UBMm5ZSyral6oFkPuy83DnBYMXgQaeceNS7bz/w4D0Xe5O2TFktoV1q3WbN55r784L7n3gNEKg3mC37yEbTspUiy7LU7yvULXtARIxhyjJvmfJsIm0x3nqwMGdvpFJaLP7EI1c8//ANWwwpV8qhOwgRhRCaJoTgnLdxHJwzwbmmCaEJZEhEDpk0fNhmjz1842OPXOsmUKYNjGWDCmlZWn7gmy9/ePjxF7GLBBsiat3HnootpVIqkUpTF6epnaUkvV7pjgBn3d8Ahx5iPRDeSim1+oNSKo/hOuesv73xyp05ukamxdrjekkkQr7nXvjvqrp6ztl6ppogYg/vuW8EkXXDGEipAHDmZ999/9WPIs/XwvpTxqgTuFPmI1MuOPXkvxIppRRjnPUkEwMxw+k5W376icc+98AVhrRRKZadm4ZItg15+Xfe91JzNMZYJw4aIrbuY6/XorsAxXqvdC8Vuk8OBUR0dNS25b57jLvx8lNUNNkhx1TZkoV8KxYsnD1rDgBskvUHDvU27e2PU2mT61o27cM5kw2Ro4/Y55QTjrZsG1pS7XonjjpJKY84ZO+rLpgkq8IkePbBpwiYy6hvjn3w/icA0LcBtT9E2B+ynZwzpdTpJxw+fsfRMppqly0FqJRCw/f2u590xTVu3Cw0Eecsnkx98fVcCAZs285eGTNl5eX6J195OgDwjKVfH1GIwBgjoivOmzR27BayvplrPJvz4AwUwFvTZ62Px/I/rdCtJsrQxNmnHA6NNcS1NrOBCFJRyPProqXJVIox3OSMNAHAqhVVy+tqwe8BqbKZDagPTzx4/ICSIqVovdWLWo9GpaSmazddfYYulbRVR/pW6F9882ttXS0ibuyRWvbHbSoCwB47jvEE/MqyoG2JiQCA88qayKrKqk0Pbziv569LljUlbcayOSsCJFBy/93HOWl3fREGYi3YgwPAXnuO227MCBVNZkNPRQp13hyPzf11ySZwJP5RCo3OYZpXVLDbrjtRfYRz3k7dEZuTdiKR3lRhdHNzQlqqnaFkjEzpL8wtHliC2MdVT4ioSOmafuTh4yFtqayLkyThEg2rw78uWNGOI9o4RcAfptGglNQ4ryjLhXQKGQOQWaYFmMY34YyOcHMCTKtDMJmIMskaG+xQ3Hr0FrrfpRSBU9aFAAq5rtmW1dwQ3gQWVvxxX505VZlGLcnm2WelIlvCpiuaxkEXRJitzlzwprpoJOIQt9TnRhoABg4sHpLnWbC6GQwOtgKpANAOA6Qji5ctB9joi+L+QIXOIDyStEZqDAEA4ww3fqe7K8kJuJ0k8fb2mUO8efHChbjXDhvo4B9YXLTN2C31hVWbleQwwXNyc1xeb2HQm5MT2nyLcgLqV+je+0aMoQKob0gBEx2hm5RlIfeg8jJYPyL2TykIAKVFBW4dU0plx5rBlpCb88K/pp9x+l95XyuWwyy5PZ57brsaEfwej2EIwcUmZin+QF0hAEzEo9//Og/8/mxKHxliPDl8+JCAz6M2OY8QkQBg+NDBJX4PWTLbIioFIuD5+ocF/3zxLc6YZZl97RAjAJQU5Bbn53o9LsHFuofr+xV6rRaaiGju3AWL5v7MQz7VipgVgeAUjh25/54ZS76pCRJRbn5o5NAhGI4ywTuoHPP7rrr20c9m/6BpuhNV7Vu1plZpH3le7yDO/7RCZ1Jt7n/4RQgMBGW3wmimCRVODCovnTBhh03AR+n06JdSCq7tvvsOFG1un+pASirmFtUyffjJ17/6+jTGmBAcEVpzvPrkBnBTbdLxhyi0ImXbNuf8lTfefeOdWSzkAUnZeEPVh485bt/CogJStOmFvlu9gqOOObBiywqzOZ7lJBAAKEXCY9RHmv928vUnnnTZjPe/lICtOV6ZokzV33nwT6DQRMqyLYZMCO296Z+ddebtls/LGFKLR8g4s5pTI7fe4pxz/+6U72ySZsRJrigvKTz77L9ScwrW8HqVVNxjUHHes699tN9fzttr7+OfeuZfK1attqXdkjyIUtqWbSulOu1Z+D8rYp110TEMvfQbWg45pgkWS8ZfePndi699OCGIu4RqSWZAhqgAws3XPXJ1eUlRh3zz9cWLvxcWX/eTnEidfsLE19/64MvvF2oFIZm22pF4ShGBNrCAbJo5Z9HML+4cUP70ruNG7LrLjmO3GTZkcFlhYSFv95iq9cs36RLjPlFoIpehO65D75GGUtV1Dc++9O7br7//5YLfwOsWuldJ2zlpkYPgPL2o8uwLjz/60L1tKQVf7yJZAsG5kxT/50PSoBT4fd6pt15w0GHnN0Timt/dsdQSQaZtQNTyc2WuqjTtl9+a9fI7X4VyjNJQ/rYjh26xdcVe48dut/VI3dDbsSVEpFQrSulX6A5qoUD3ff/jYu7PDTfH+DrqNAKlpbIsKxVbXdO0elXlx3PmVy9eVttkgkuI3AA5fYAAAYhpgixIV9Wfe9Gx9955GVFrh9L1fDjW2BxtamqKROOt3eI2GJzKVE6XlRSt22uPjKGUctx2Yx67/8rjTrnZ1C3N0GVWNmkLz0bOP6JgvDhEAGHTDlfVzlv0G7yBtxuYk+/batttDxy/bcWIzUcOLSspLdYYB86hpXnpGr11+7ap8cal0IjKlqwo/9oHX4J7X+jB7nKEWBpSNqAEQNB08LvAcGmlPkBst23I7GiMR6I3XH3W1ZefyjN00vout5IKc/zP//vjzz7+zsqOX2wwpEFpO5Dnnv6fx3wezzqWrHPOpbSPOHz/f/n9J0y6NtLcrBUGSColVSf6RuSQm5wz9LtYyEtKxS0ZT8pV02ZNe+MDLMgdObB48/KyHXcZPnLUyHHbjMzLCWafkM5b1A85AICY22BZRmOdvJCAFxkCMAICUiQJiJzqo3aGPJ4cVBCa+vDVEw/b34G6fXRQEnIMx1NzGpt/D2vEEWJmqZXb02/inCulDttnl/enPXLShbf8MudXCPiF1yDLpm4cA0mO78EZgiGwJIQsR5r23GWr5y5c/u/PvuSglebljN2qYv99tt9jt52HVgxyzg2lFCIh8v9xhQYgaPEHVfsTkTr7fwAAsGVWc/m2zmttV0BE07701EOvu/Icj9stpWN+1n+tW7rbE3FNoCF6/rc9/0qGkmsev7sXr45DWWy/zRZfTXv0rgeefeKZd1Yvr4GSXE3jQA4BTV17NwBAZEmnc4MIuBE9SpKt1PJIePn02a+//p63MP+QXcfusc/OBx+0Z1lh3qZtrdcdWfaKJUBs+Wm9CGVbGi7YwqWVTz776oLFKzjnnHMpVbfNmXp2q5maZUlKkpKqs58Ov5Jd/6rDX3X676pXrx86dtrn8dxw+Vkz3rz/pEkHeqJRq7LeiqaUAi4468QHcLK7qOUaSERKKmkrIMUYCpch8gJiSEWCuV5+57MzL52674FnXXvLg/MWLXeIvzXIbPqfUuhOHzj7X1RPXwYEsBHf/PDb8694aOcDTj3zvBt/W7qCcwbA+qIOiNr/P3X9sd79qo9BuENOK6VGDt/s6Udv+uyDpy4577itygq9yrJqmuxwkhBQMM454xwzTR6oK6udoVmlrcwU00gfkM9yvPOqaibf9fzYCSddduU9K1bXOOVt1G5uy/8KbQdc48hYS5IuZmXrrgk5Wv+dyFl359SkzotPuNcNfl+jZT/2xJvT3vniiutOO/ukI5wS/PXPs0NExvF34KCRIQm2foc4IYLTfxURt91qxLZbjbjyshO//u7nTz79/tPZc77+aZlsiMmgBkIIv08IoZSt7LUlFSESgW3aiCC8bvB7UmlrypRn/jPtizunXHDYvuOdXlqITtdC9r+h0IJZDc0Qt4AxINWizZSlxE69CWuPQRkgA0Ggi5bedpxzdMYHQWtPMiKQNufIK0pXhuPnXHL3zz8vuOvWS7wulyLJ1sd9QZRpSyad9ocbWKkZQtxOBd3rYeZaslkc741ISZUbytl/r93232u3xkjzb0tWzZr9wxvvfr5q4YolNfV2PAFuHwTcqHEuWIu32GXFGhEQSVCAgunDBy6sqv7biddcfskJ1150MmQ6CbL/AQtNgJypxsT+48cOKi6wJSJIIOdVVu0GMZENIACwpZIKAVkibafSiUgiUt8UTccTqxubm5vSAAReg7s9CEQyk5ZERHba0vwu5XM9+sRby5fUvPDMzTmhUK9bNiKismVxQWhUaWEqbeMGdoAYopm0cov8jlqsP1HDEJngSmW6z+QGA7nbjNh+mxHnn/m3VTW133z107xflsz+adHSFUsXVTakwyZwcAyHbujkEKNOQ93OVNtKmlquP2HK6258rPK3lXdNudjr9m4aXYzXaqEJOadI7dmnHHLIQb3pbUdAlmlallXXEFmyfPWiJSsX/Lpo2vTZC35bCUGf8LmUabf2WJRSIaA+IP/d/356/kW3PfvkbYjYk4Vu7akFyDg1hk84ZeLNl59hWRYXG5aocg5spZTLMPowcsFaqloykE0RIg4oKhxw2N4TDwMFFG5s+mbOr3PnLZ71w7zvv19cv7IuVt8EnPP8fDQ4SJNUZzYbQdqScWTFOY899ppS8rGHb3SmR/V5fe6fEnIwsCzL6b/P2pA0dcKUtUfVCISIum7ouuH1+gaXl00Yvz0AXHB+3auvTrvrnmerqxu1ojwprdY/IiBpW9rQshfenLnVdi9ces4k5yBdty7P2TdGQKALoWma2IClp7+TZLI9OVJLwh0icS5yc3P3m7DLfhN2+YdtNTaF5y1a/sFHs76d/euMr7+DSBiCxeDRuKYhI2V1DNY4jqMYOvCJZ6aN2mqzf5x1QutEyU2fh27JLsoe3ImdQsDsf6fMVEzKJk0Rsbyk4OJ/nHDIwXtdeeWdb7w3WxSGCNtqU4iQpGRB3613PLXLdqN23nFbpVTPMQODtoZ6si/o7XUlKjeohUNwhlHwNhKaFABqQisqKCgqKNhz57GmUkuWLP1y1g/vfDBr9vcLKquaQBMix7tmsIaIUElWHLrulidHjRw+YbdxvepivLHSdr1afWzLKGfM6baGSikp5eYVA/7v6duOOWRXuy6MjGfrgZKKubVwJHXL/S+qXmI7av8q/k7y+5ptQATGkGU6jpKzsDrD4UM3O2nSka8/d8/0Nx+644oTR5eX2pX1kmDNnBYlFXeJSNSa+uDLsPFn6v3Ojm2bI885l1IGvN4HH7xmm203t8PxrNgBAYKybCwMzZj+8bffz0XEnqe0/1Fhgt4pRK9zmp3wCiI6cyMY55wIWhsZjxw2+LJLT5/+34cu+8cxorHSiqWR8w7ZB9KSvCj43vtfvDvjc0SUUvUrdG+Ecy6Vyg/l3HvLhSFDrJlszRmTFnvptenQNlC5X9YJH7Y2MnYAd0lB3h2TL3nnjSeH5Pllc4oJ1v71Qc65TTD17sfpj7QFG7lCOypLRLvttM1eu4xRtU1cE9mLiUDkdc3+bl4imRSCbxplya0FBy0IQTnj9jZQk0SnwsUx2PvtufPrL04ZUhSwYykmsreelJQY8MxfUbtsxUrH0PQrdC9FSgmAEw/aCZJhRNEhEwN0fUlV7Yrlq1rHkm8C5jPjUWQQAuPcace/AcGrY7Aty95m1PBbrjnXCMc7lNMTKXQZNZHYwvlLW/33TZnl2IAbzBAAxozdJq+4vDEW5y7RCpeJADmz4rHlKyqHDx+2CSQbmJYVjydiSTNt2qlEfPXq2ngsmkwmZ30/75D99txvr51aSPcNkoMvhLClPObIfZ5/bca7Mz7TC3Nt227luZGBKdVPC5fut+9uG2/Pqj+BQiMDoEHlZYOLQw3LqtAdyOraSMggacrKxtjGDzQAEWsaImf+44afF6xsqgsrs9k0uVQCOVfNK3U09ttrJ1KEHDfYUgMCMsCzTjhw+lufWlJmT1tjgFLCj/OWAgDbaLmOP4FCA5BSPo8nlOOC39qPjSMgRJOYlU5t5IxSxui6dW3RwhUra5rA5wJ3vjNUU+iaXGHVVVURSWScaAM+qQNsttl6ZF55YV0qgZpoQxcMQCqzqWGjzlL6M9w3OjkHgncJ3WwrDRu3ECIQUV5OoKwkhJrQGOfIGAEokqatQoGZPyyMNMcRfw+Swet1V2xWDCnZsXaTSNvIM5T+FLePiIrINCW0X18nkM2JuDBgU+h8TojoycmjpEUIBI6vgCQV8xorFyz8+cdfW7I2NjgACrhdoJw5v+1OA1LZ01CpX6F7sbgKEevCkar6JAgtuzsjIqAinaPH593YuQ2ATFr29qO3wHgcWOs0KgVISBLd/pff+gAAfx8laozHof1ELGdQuXL7s7QC+xW6x+IEtObPXbx06Wru92TzsQgoFeiGPnJIMWwqstN2oygzsaDdKlDIP+2j2ZXVNaw3YdF1tR6O0U0l08uW1YPuUPstnJICEFgxuAQy0G+j9Fj+eIV2SNkvPpudjkWFy2gXYEAA2/J5fHmFhZuEMiMAbDVqaFnFABmNMc6zRugSN7SVqxteeX1GC21HG8h8AMCsb39sWlbHMiMEMulUCgBQblFe5Byb/Ri6d3gDGGNN4fCL0z6H4iLbNLOTFxljkDYryvKLCvI2gfRzp4loSWnxIfvvqMKRDo10EYC4ePShVxYsWMwYk31spJ3EPKdUyH76hXeUiwRjbfXIDMiy/Lp/5OgtWxSa+hW6ZwcfACllI+LTz73+y9zfuN/b4pG0ZrUjpq3ddtna7XJJJTeBnGbH4Zt45EG6EbBMu90IXalYwLeoqu70C+9sTqQ4Y0rJlrVS67HObTEapZTg4o1pn30840uRF7TbmswjZxxiiREDfVttPYKInPTUfoXukYdEti05F19/9/MdU19kQT+2T/JCAFtJZGriwXtBy6S9jV2c0u4dthu57Q7DKJrE7IdCVKYlinI+nTXnvItvsWyLMS6l6iM4q5yk8MZIZMrUZ1KGCzi2JthmPNNE8qBD99WFkHIjrlv5Q/pDk7NkQohvv/3x2L+eX5cymUtT1BYgBCLUBNWFDz5gz61HDycitkn0HXSMdMjvv+jUI3Wy28MKAiQlFS/Iee61Dw4+5pLflq7gnK9HPqejlE72OXDOV62uOu74K7/++TcR9GS1ECFkZCfT5RUFZ55yDPRNr58/vUI79Vfr8ZMRqRRD5JyZZvq5l9468IhLl0ZNEfKr9k0KkXGVsgM5wSsu+Dsi9hrSrfdtr8/zUjc6fdghE/beeYyqaRK66HjHpHhucMaHs/Y/+MzX/zPDsi3e0nmx+y5KHdFGpgOydFKgvvnhl4lHXTr942+0HK9qPzKPIcem6OknHltQkK+I/gdKsIgMQ2MtlT+9dvAdWbl69aefzH7hPzNnfPKN0jUt6Jem1a6QjZAJZjVGDjtyrx3Gju49pCPSNMEYMvb7R/ixGyMtpdKFdsWVp3888wwzYaKuUftsb5JSL81f3Jg45pSbDjr4o5OP3uewA/ds7VJCihQQruXr0YmrA8CSpcufe/ntx559t7q2SSvPl6ad1dcBueBWc3yrrSvOPOGwjJ+Om3aRLBFort+WVs2fvzSVSvWs0S0CEMRi6fpwk2Umv/5hwSdf/7RyfmVNc0TpnAV8HEHadrvdJ2CCWfFEYXHetZec5CRM984lQsEbGsKLF61ojsbE73mMIgAR18VmFeUuQ1/z95wzpeT4cdvcN/WKM869nZflArSNMXDETtsioEupvfXe59NnfFFR8vikv+1z6L57VGxe7tJ1vg4At6ahYcHc3558bfqMtz6pSSXB5RKFAZluXW0EIIagTFskrVuuPy8vP18RbexD9Nalna7NSkvPv/lJuPS+HjsKCKAQiAFJABNQQNALhs7yAhpDJSWt0V2NcUa24o3x++65ZNiwzZSiniQKt1bvkpKK5YceePbtB+5/5fc2OZxBLDFoiwGff/TMgJJi1YkDQIhMKTrtxMPnL1p2zwMviZJ8INWBg1eSGAILek2pfq2uufKaR668/qk9dhw1bMvBw0ZVjK4oLSsp0t0ety4YoiKKJtKxaHNtXcOXPy6qXVw1a+6vv34/D9xeCPl4MIBAqq2hOmXAD6Ksqb/+utMP3nvXPmlVtZFADiW53wUBdwvsXnfav9VdbulwIAkICKgTR4eAaUxailY23DD57GOPOlgp2UOw0a6/ApFkPjcGPL8/lSHj/mBe0KtrXYAPdHKVFKlbrj934eJV77z6vjZ0gEKi9t1hiEBKxRDR64KAW0n5yXdzP/nqB/Dofo/bICZ0Q/doSEBAqbQtbTNpWolEEpI2BDy8YgCSIklOvy/IanmOHAGYrK4/89yjr7nkNNrIi717qNCZnJWW3IOe4dg2kwCAXfXyQca4LsymmNuyrr3jrCsvPHU9mlNl9QwhIvl7BwgIQNm2Mm21Fr1HAnDp+r9fvuvSwXdMfeItDHk1XdhrDDknAiIFUgGiVhREQJIqaskoSUjFINl6RQ4IIISWF0KOSiplWi0uXrt157pmJ01qaLjyor9df835grNNo20S/B750Nipk4ROIxhwqvCFbqctc1n1wCGlU24855gjD1CkEHCTHxGCgEpJzvCeO67ccvTwC654INHUpBXlEEH2aN2WnrkIANKSLc7cmqMy0clIlFK29GNjWWepM2gTmWBmfdhj4/U3nn3pRSfjJtTYDv6QBH9krZUTqKRtx9MQbtT9xmlnHP6P8ydtPricMrgZ18NEbjy8KeNOL6TTjp+41ajhV111z0effgv+IPcbDi5pGdyGaxCSa2aPd9nvGBEQOSHY8RTUNY3ebuS9d1wwYfwOAIoANxlt7qjQjDEuBNcE9mkOY/bQGmlJ27TAtsGSEE2Dz9h2+IDx2299+F8P2W3sKABowc1q3RUakXGucY2TFH8GyolxRppgnK/zCcMYkFJy3Jgt333rkRdefPupZ1/9dt5KM2WCz6153MiZtG1SCoi6hG2dOxTIBOecpZMWJONgy5HDSo8//29nnHFsTsDvNEn6HY5BLgRHbAtNEnCNKSE2hA/aTqFTyaSsr5NKgt13TTCUApLgtOWXEgLeoNBzgoGhIwbsOW7k8JHDx++8TUF+HgDYts156+S4HjyqZSVkuiq52gfp1J8iYMsQUul4ULMsa93BB0OmFOlCnDxp4lET95nx8ew33pz+8cyfqmobwUxDKMj9bkQEUo5eUMZ7zCqSb4V3GdjMiMhuaLbjSXcgOHTE0LP+tu9hB+5ZWlIMAL8bp0FETbX1UqHEVo0GS2PQFE6m4n1fCNyyBASAm1UMmHDAToJ7+6pJBCL6vYbbEG6vO8efE8zx55XlbTOqoqyspKQwv5XSJkWKqBfTBJ2dK68o3nP/8YZWLKX9p0DGDJWyyspC0AONwYyhJlKK/H7fEYfudcShey1eUfnB+1/8+M3P076Zu2JxFaRToLnAxQAZIABjwLFlIQiIQBEoAkkgFaQtMPh2W22+7847TDhg3LjttvJ7XAAgpXJ6sv0+3pPmck/YfQdlUbavjoxLFdt6m81Jqb6tMG/X7MK0bOprQKV3zbu1TtFbv1OPbEmSEJn8E+XTEJAiQxO9tmoO7dC6MnWN4apVtctWVi9YtLS2rmZlXRhSyVg83ZSymJKNkYhL17xef45bePy+nLxQSUHOqBEjhlYMGDyk1O/xtOHu373W2HYIWuyE2QKpDI33LUr8Xbq3UIZzzdTyt+hwX41v64xF+VMJ9brhnVJSETDknUaX4olkwrRBqXQyzjVNaC6PS/e6jQ4fc8aL/UGDZdf67H1cYd5OoYk2EEWwwVfyz1c/i05JbB8+YPburDEcNvuTzmpQ6wf/cPaz693BFvoBN4hCb4TyZzbPfX5v1C7Wl9X0IfsXGxt536/QfXOa9z/spikbO6P+P7XB/dq86St0v/RLv0L3S79C90u/9Ct0v/RLv0L3S7/0K3S/9Eu/QvdLv0L3S7/0K3S/9Eu/QvdLv/RGNlBN4cY+j7RfNh5pn4y1IRRaATgVQqw//aBfNgELjU4hJwFI0+pX6X7ZUCCAMVSSMYZZxXtig3wTkBVLLL/8CliyFD0eUKp//fulr70/phIxtvkWg26/VQjR2mNyQyg0AWOrH3oo+siDbuaHzvpH9Eu/rCduRgIbYsEDDxQeNyjVWo/cxwn+ThFv86/zF++1l7cxCR4/KNm//v3Sx8I5xqJy2xHDPpwh3G7KqtPpUwtNhECKoOryq/SqagyVkZXu9wv7pe/hBlHUjBZefJHweDo0QugThc6UBpFSyPmqZ59PvPOu318spdWvzRv1sZ75aUWS64MeETKsF6731YQww7W+vxxadsREWKNpat9ZaCJkLLmqMnLH3V6mSYYg17ZevejLT6qji8lYu4HKSq59mTgDZNBhvomUHXeAd0U7EiI6TYEh055r3bZ0LU3XyWkK3YNtRga96NKy1q/gDJAxW6JlSsskRYiAQgPDICFAyp4pIuPAkElFpgW26TRjYUKg7lK9uBoiptIyP6fs8iudWSUd+sj0iUI7DSwlMrF6yt3461wIlED3TYwQ0bZVMkw9i74QYwZ6vO3+KhZTymz7gMsPmtblGiEiMohEFZhZn1AAHDz+tmZCiKCkioQ7DQ4hoNPcU4ABugFuF61dCxFAUaTL53WYTgJiLh8Zxjq9J4xBOknpZA+tHDHNCy6jk+sjOkQYNMeIzBTwdMCrDywx3G7TNGVTs2iMcivN3UHSBCh77ccv50ik4hGQMg0IAb9VnOf2+UipWCTCq5oMJbnuVS5j3R0t5MyMhl3nXhjccSxkmqa2qx3uG6eQpETOq1/9d9VxR/ncBQqgu/1giMm0OaTMdeiBzFG+dam8JwIhrF8X2u9+oHFBAEjK4sD3maAP3xxsGwhA19JvvsPn/0ZuF3QyHEAgqXS0ju2yu3vXnZyoP0nJdD1l2zRthpi/hAwdEFnKTOX7vcceAS43KNXh9hDRsky7OaqtWG3/PE+t+k148502Xl1rszR17j55EtP1Tj8mlYKGsFy62Jw7j9dUab4CQux2DRnE48lxW3n33JNLua7Nh4hQ09Izv4CvvkHDnXV9AsYRQDY3WW6hjRxj7DLOv/suqnywp6TIcLvNtGk2NcLCxZFpM2IvvySSaeYOgrK7PGoYR2QUaUxz4KO29I7b0bfnbmLYUMrL8/j8iigRbpLz5ydnfBJ56SU90oyewDpxu4xhMh0vDm7+xUzvwDJQtOYB1RcW2gEb4eYVk28Lkq4Y7948M+BpK5x31dSiEyatO8R21n7B8ZNUKgL+AlASAWUikXfUEUVHH9H6sYXLl1lzv9NYGbVH8Ch0SMbS6Wjg4gsLr77SyMlpuRkwbXvF5NvTtbWgZVbHNMOhU/9RduO10BXmcAx7Kp1csKj2iSeSTz6lCy8x3qkKMs4TkWr/lVcNvnVy9xeUiUR87rz6p/8v9eRTuiuHWBfpA8jQNtMFOQOn3p2zw/bYkwU0Y/HlH38KNoEr64XhGqZT8VS9MX63/PMvzNltF6OgIFtTdAAoLoQth+ccdnDtgftUnXGmOxwDo/MIA3INzHQqWa+N3bHgH+fm7TVBKy3ucJPugjwYNpQOObj6Lwc2nnMuW1zJvGsnxBBZ2moInXOFd+AAULJTuCXWX5udr6p/9DH9p+9YqFTZZnczoDQt3VStHXhw0XF/ZcoGYK3tNLv7DimZ4aq8/3Hrn8+7giUkpYM/UCpzdQ2SctocAwDYdsd9RMY4t8N1siSv4NapxSeeAADgWDXGot/NqTz3PPrqK83IIV0gYxSN4I7jis87mxFlNqyjmiIiAGPMZfi3HuV78L6lfm/y9ns1fw7hGh8WQjY36NvtMPCKy1grkFhD79Fpn+jxhHYYGxi73UqXEbnvXleglKAT7IGIdiIWvPqyvB22zzzIupg3Uqjpq6+5SX7+uQgWUIvPgEJAOGwVhQqmPJh/wokuvxcAQNqkFCBzCISW0d+IiEWHHSLjsfrTznbbSvEOrxyh0GW00Xa5QjffXHzOOXpOyFltUgpaGsBlGjsRIGMle03AW29tmnQaStnlC9yyknZzPY0cNeD0U0E5DcSwM3O5vvqsADHy7Zymu+/2eEJKdgutOFeJZntQWdn99zBNIyBABC5AdP+joeGKL17ScOtNuh4kaqHQGVcqnmysybTiBADGeEG+bGctBNq2GV4F++1Z/sGHxSeeACRJ2sC5Alj56OOL9tobvvqGhQpJF0AEtrQ0Krllspaf50SIgPM17oe3uHdE0kaA0gsvkhVlKhlf8+hHy6ZQqPTuu7RAIAMVOrmgyCAWIrJtxrDoqivlkAqViHVyQSFkcwPsvH3xeWeBUsAQGFvbAgoQAjW94cOPYk88zN3BlhHUgIyb4WprxzHF77w14NxzXH4v2TYoBVygpqMQkjFiDLlALpBzh8vKO+AAtuXmMhmB9vPykGsyXEMjhw6Y/nb5NVfrOSFy2lpzjpqGQkjOFWPIOXKBQgAiKZV/wH729ltTIta9g8ssKQUruP4GLSdEANDFKCmxvmADkJSqvv56rb6JQoVrGMh2SJJJSqWTOZfd4dtsMyKJyAGxed58lUx0g0ERQEpZc8/9rpoIhUKtX0EADEA0NWUcGufPfYG29m9cU9HmlEsFL7uq9NqrDJ+XbAsQkYtUTd3Kiy9N/usVD3eznEKyLCBiupFsWuE57YycCXuQksgQABPLVpj1dcDb4ATTNK2wyF1U4CgEKeUqzMeywfaSVTr6sp8BdSPduNJ/2kXB3XclJREBkDX9thSbGrMvyDXdVT5QCwYAADkHknpBvm/kNnLpO4C+NXx803bpA26+Uff7SEpERlJG5/2qLAu6ht3ImEynqy68TE8p9BskbQBCrpmRGtprj4pnn/OUlZBtAaCTFxGZ81N01lfy51/T4VrdHWRbjSw44W/unBARIWNGTsi9xRbp734SbR2fCbmmoo0wduvyZ5/zjRgOtkWYuVps7rymL2fZP82XDXVcN9ioEcUnTTLyckkpBCncbjl8mPzkIwEhgi5ayHIhI3XigP0KDzoAuu3Au74KDYxVPfBQeto7rkAx2d0yG1zIcK3Yd5+SEyYBERACY01fzFp+2F885CXoGnggKiUhngCfr8MLwwBVPK4IGDIgBQDuspK4gzIBVbg2XlpQ+vBDJYcdmvFchQDApo9nVl9yOXz/vcuXR4yTZQEACk1GGmjrMaU33dhCkfLk8pUL9t/HqG3k6IOWY5dcLOLXhj76WN6ee2QNj5IdO+tzTtEmtVlFwQXnZTo3CxFe9NvC/SaGmhqAtVIxCAaP5OUMfeTe3PG7OmECRFRaJ0cqMpZK1XouvSxnwp4ZChZxxcNPhK+/1sU9oLAbeoeUFMk48wTJ8XCEgHATjBxd8eRTnrISkhYyBozH5v6y8ubJyQ8/NMIpkDYDO4EiTSky0+WXXOAY6cgPP6WXLme6u62POBMqFrEGlw5+7nnflluAlMAQmUgsXFR16+2RadP0xmZQyMmyQYtDnCINg2+8ETkH4NGFi/TfVnLmo67GNiAyy0p5PKUXX8A9HnCQap8qtAJAUACI8aXLaq662itCa5lsyhimUqlAcNDtk3Wvm2wLGU9HYysvutTVEAUvg27HWXMEMHTqjEFTlk2KgKND3shAALlAKZPROvcxR295yy2ezQaDlESEQpBpV95+W/2tt7psjYWKSdoZRwQBLcsUbODkWzzFRY7lU1KuvOJKbcFC4SsmO95K0fCY7Y6kmcvd+konq2rk6pWa0NvdGkmTq7I77/IOKifLRs6UVPVTpniX/ozeYpKJNs4nZvH6SlBt8NpsbErP/VHnOmajSiFkNMx33mXQ1Vc4g4+Ri+jCRXUXXxmwidwpUN0gUEJEMFwZsIGIpmX62ICH7/MOLgfbBsaBsaqXXll95llGc7PXk4sen8PxI2fYUKXm/IyMRVdW1j70cOMd9/oFR7fjxiEgorRNDUqm3OPfcguyLeQckDdM/6Dy1JPFqiqvKwfdAeAcEIFzX4OMfv0DMJaurV91zwORxx72NKfQFyQpu6DqRCq6yv33Sbl77UlyLQ3Se6fQDIgISQGuuuV2PZEGf373zAYCplMNwZunBLcZQ1I6PlDNY0+wr79lgQIi1T1+ok55QAfwJJKgVGvMAhGUtJrN5sK77yw791ymayQlIiAX8aXLKy+9OPX6625PPrpcZFtt1CzX09FK98mnhfbfD5RCJGB89cv/Sv7rZU9wgCIFGrbgDSOdWOb+yzGBncY5jg4gNn06ky+r5i6fylgsQi7s5hrX6aflTTy0FenWvvJG5JlnAoFS1QYBCTRdJZpdRx4W2HnHVoPU/OmnbOkK7g6pVscfkdlWXJnl114jggGS0ulEW3X1dR4rDrnFYNmw1jhVi8YzLpLRKvdpp4d22xWkJMaQsfq3p1WddIofXRAqJWmT4zqDJMVQeOzlK1c99EjkoYflr7+EPAVKaBltBkDO7Gi9cdQR+YcfQrblvAbN332/6qQTjZoIyykj22y5GoCUoLtVfePyR59KPv5Y+odvfJ4C8gWpK4oDkcyknZOTd8EFSERAgLzPFRqAFDLe/MFHqddedhn+tYyw4FxFwtrOe5adfjqSIgRgPPrT3PDUuwzDT0iQPW2yJ7Q4AVc1q1kqBZoXSAEwwxdIjhhRcdfdRQfsD0BkWyg0AFj50stNV13Lly1zhQaQVNmkHnIhk82qoKz86suZ4A6nnqqsiky+3Y1+Ba2xSUKhyVhjumRAxe23cQACQsbMVKrpyf/Tpa2EyLzVXEC8mYaPKrn2akR0MF+qrq7q6ss9SijElqgkodAp3mwW5gy+9hqhaSQVcmZF4zVT7uEKFG8lQIkJPdG0wnXIUfn775N5kRirfvFl8+13DW++su12C7eWIA/KVEKVDiy+6GJwZrpxHl28pPrCi3ygkcsDltnOGVUKXS714y+RWZ9zTdMCpYpUW2AFAdMy6fVudtEFGUqR8XR17erTz9ZrwyyUT+lU+0HuBLrhXrAkedYFyJgnZ6Cy7a7DNMSQpxPhvCtuyNl2DCjCtY247hXLQRIQ03X1y885x4gT6HoLvuwCbKRTdp6n7O5btYCfSCFjUqkVZ55prG5iXh8jYJwzlvlZ94kbRIqjoWqbzEQ8w3wBBHbYfuv3Pyg6YH9SkiwThWaFm5dfcVXd8ScYlXVaqIRsGygLqyGCUjErmjflZqNiiOM1EuGqm27DX34Enw9smRkkrBkqEbW97vLnnvFUDCFFCASItU89I2fOZL7c1jMKFaVUKveay7xlpSRtQiTEZRdcJH5bygO5GTcAkekulYgkNFn69FOBrUaTUoAEyKqff8Ga9anw55G0WhGqjIVhyy0rnngQiIAUMpaorGy88RZDARgaA9a6gIytZfoWMmanI94Je/m22LzF64bEJ5+nflugpEIg1I1ODkwELZCH7mBLWDTbHDR59tgtsN12LReD+C+/JL6fzSUHK810g9ZgJAgRAwHw+ZRlZV9tTbAhYxEaMbrwlFNbpyStJdGjV9aZkPPqqVP5wvksOJBkt8QzMjPV5D3nMv+O40hJIFSKFk++Lf7lFz7QzcaV2aM5yQn3BUK4TiUBhJpGTWmy7NYLuIqLACAD4zQj8t33qy++xJ75qd9XoBinjhw5IdfscKX3mKOL/3qcExRExho/+DjxyvMuXyEhgBBICGRZTdVWKFD21OMFe+9FSjpOXuPs78K3TDaES7FM6RkKzQpXuyYeXnLcsaAkISJj9R98kH7rXZ+/WBKh0IAUps1EtEbl55c8+lj+QQeAkqAIhYh8Nyd8251eb75SKuNlIqKSJplF19/kKioEaQOilUguO+vc9MK5BAY1RtuvoWLoFwFfl24JIZBtDCl3xnQiZwDgP3i//Mm3xV95OfnzjxqgLgLgpBhk4fIuzmFUYBlF+Sh4C0UP/h3H5U69P/78c9YPPzKSHi2k3G4iamf41oU7VyRR5px/lqu0qKtISse76XHoWylgrOmzWasPmainSela13dGyDUZbbS3Gbn5h+8bgYAT5pDSjkz/EBqbVIc4MBEiplasCk++SdgacLaWfAYixkU4Fd3y+y99I0dkQuhEJC0UOgBUPfZE7bVX6/UxHsjrPN2CcUzF46V5Q2dM9w/bjJRExlNVVfN22Nm9ajnoQWZaBGRDSvlzXfvsXnz11YFtt8kkWnCeXF21eO9DxYJ5PJibeakYp2Q8XZa72YzpvqFDiWxkPLlk2YLx48XqSgCR0XrNI3Pz3QfvXXzuuYExW4Mz9liIZGXVoiMO12f/hMEcaAt8aHa4Gv7+9y2eeIRpwpmWZUajsf++C1Iprf0aKgLOIl9/nbz/Ad0V7HxzuZCRGmPSpM2efZKInNkWzm+sSHPze+9Gv/0u+vo72tJlTPeS0e3x6xjRSDWfeOjQ115l2dl5AOl4IvbpzMhHH4XffE9fuli43CjcPShfEoIijamKzUbO+VLzuJxJrX2t0EQAYMUSC484ir3/EQ8WUDe+IGPMspLKGvDW67n7TsguK+hGFp9xRuqJZ/Rg4VpIQAAUglKxCKWGfzErsN02oAgymSqYrqmrvPra6FOPu1whdPmy/L/2jirwVLQ298nHB5xyIkkbkCPD6qefrX/ocVFUSNJShqHn5+ubbxbcZXxw/E7ocH8AwLnZFP715JP0f78tgqXKOaOcMop0PPjkQ2WTjgelCAAZq/r0i+g7/9V8PnKSHzTdNWhgYNttfSO2BACSJgIDLhIrKxcfewyf9ZUWLKaW+BQyTqlEOuAd9v1X7oFlQCozwa1rrCgJ5o/fnX/5PfqD0ClvwBikUmZBoGLmR97NhpBtIWcAzOGYnY8kl6+q/eeLsTumCtMGXe9WCxHITrvF4PffC4zZCmTGjjq5xM4nEnX1jc8+H54yhTcnmO5dJ51GRKKUGS5+4cXCoyauo/L0WKFJ2sjF8gcfiV50qe4JklzLu5uOrPb84x9Dpt7DQAFiKwqiNZIGSUpmGHVvvFVz3PG65iUuujMMiMg4ReoTQVfR/Q8VH3sUasIZJAxI4S++XnXJJfDVLMNfqpBanfE13wc7XEMH7Dv8zTe4xgAz88XscCTDmRCBENxlZLaZFNk2agKAxZcsW37mGfD+xyJY1PpKoxAy3CT3n7Dlm68yXcNMBnB3Zx1JiZoGANGf56449XT4+jstWNQGnRERIJmoy3/0sbJTTyYlEZljU2nNID8RKQUu18ILL7bvm2oEB6iubQ0yLpsb1X67D/3nP4283MzNOO6kIsRMpuvqZ55rOPNsQ/OvxcnkQkbrceexg196xTWgdI2rAXABAJWvv1F7yhk+kysNYW0qjUJY4Vp25MThL/8TGVv30W89UGiH84/+On/ZvvsatTFydXt8ME6pZnPLLYZMe9tfWuxM3ev++vHlK5buf4i+cDn5fSBlN5wJs+xEoknstXv5vVP9o0a2nR6IdtpcePjhNO2/eu4gZaW79PcZw0TSLAoNfu9t36iRGcvXMjiqE0UBQi4AwJaq4dVXay+8WFQ38WCo7ZVGBspKQmrIl18GR4/M8judvMCOeegOtwAAlpmueeyJxhtu0hsTLBDKZq9QaFa4Ujv62KEvPp8Z67a2NWya+dnKAw42hAuArwWwIcpoc2qL4UVXnpd3wIGuwoLsZSSlkBRwseAvR9Fb/8VAbnc74rit0XrafLOc88/L+8tfjJKibIBKSqGSILSFZ5xtPv64kVOqup+xyxia6VRAH/jGGzm77Nij4Zw9cAoRCGxZc9NNbNUqyCkHK91dlJsoJa0Bt93sLy0mZSNyZZqNH3wim8PIOqa0AAH3ehtffo0tWAj+vC7XDoFxTUbDKUMLXn/twMsv5W6DbAuQIc9cE3WN+/ySuZWS0O0wMYkq/5orfaNGtmgbdW5QETOal0jGvvm25v4HUm+97WIGC+a2wS1EREglG/Puujs42rlg24DiTEB7jXh+qrGp+fMvwo88kXpvumF4MZjTDr8xppJRVrF52S03M85A2cCEjMaa3p0ugdacZM8YVwgNUx9wW0zpovtAlfM3IhDyLF1Ud+IpTSO2CB1+pHvf/YLbbSM8bgBChqAYEPgO3L/hP2+4YG3D35UtAnmwuLL+7LMbHngo58AD3PvtG9h+rBYKAhBmdJpCB+xb/fQzYNuA3XGLiMxONXvPuShnlx0zJ8Y6yzortFLAePVrb6TfmiZ8hWSZ3Z5BmhmuNI48Onf//UApIASG1W9NqzpxkqE6D2cppTQphC+vC4KdgHGmyAxXpweWlD3wQNFhhzpZeA7NnJVYCW5/Tlx1V96DnNvhBnvczoUnTgIiR/lsRdJKZ+s0AkAqqeobw7/9Fv/mO2vmZ3L2VywS9fhyiPFs5UPOVbheP+wvZWedBUSIBMhVNGZX16ImAIBsm4WCIj+v1XWrfPb58GOPq19+1uKW25evGHXwRhgXsXQ477zrvEMrSNrOpq6674HwzTcxrRNHnBDAIgPdyuNeuzZnjlyJhs/rCcoFK5smT2566OGa7bYfeNtk/9htiZQzSFy4XQRsXYpKSErwetzMZy9eHr77ruYnnqgbs13RzdfnjN+lLSaquwgkZQBx11F6O5306BXnntVGrfaxQisFiKna2vrJt+gpIr/o7gBiXMWa1JYjSm6/xUkhQMbNqurqiy73JST6Ap0ibxRAOifq4rJCh1QibZnqoAO3eOBe75DBDmuBXKTqG8I/zyveczwRIChAZgYCNgOjqwVDpHTK9IohD97BXIYTtjSbmuZddGl6/nzUtdZMNA2YKxJLzF2oU0IDgcCFxwXBAlKyXeYuokolVHnpgGuvYS7duaCMROfuty/M/p4VFDDGEjUrSx96qOzss8m2UQg7nW588UU++wvDX0IBVFJ2wJQoNLOpyth3v5LTzsiQZYw3fPp57K573cxHwDu1mKgjMdZ5AJxxoM6qYEgpW6IvYEAALZn+4N26LYb6x26LgE7EwKqr5dAx6oGMU6dXU0opydx+7g2AbZufftD4z81Du+6CSESERIklS23bRK/oMhKHiMhSica8O+5wDR7kKE+PSLh1UGgip0y88vqb4ec5ECqFbvkHJJnWsOC66/ybVZCTrYu4/Jrr9JWLMaeMbLvT0rrOg9vOI2pCNdWlg96cKbeXnHmG5gTzEIHz6IJFK045owFV8WefkFIOpcpDAUYMSHVyriEwYOlkY+5Nk4Njt2vN76n+v+f4M0/lMC+1q8JgjBkhl640HyEBARCtGeFnXEtFq4KnXBDYbgwo27H3lY88TrO/8vmK7YTJdcMGcAVDjgIBAKXTIm6hHiLGOjENjEEsJgeUDbljiub1ZF655ujqa68RkRiFuoSzna8hMsaYHQuj0FHvrPIKMBP9Rky7/IUTJwIASQUMpS3Db7zB0NNSOUaAHBmTsQhqAoWrM6OBoCQpYFwAc/sP3B8ZkpSMc4Voz/zIBZwY7+oRkHHZXE+jRhafcYaT2tDzrIx1inKzyMzP4o8+aXgLqVvngGmGFa1xHXxA0ZETW+IUfPU70yJPP+3yFSrbbiv3XfOnM7uCDNJNK+1ttyr795vl556tCSTbRs6BsdqXXlt20MH0xceFhpahBAgQgLsNoC7sM+NWtF7svnvpReeDUs7Z1/zLL+Epd7u0XArmQrAg6ydP+Ty2EKQkSNV5nR8XdqQGdtix6JxzyRlmzkT0p7nh+x/06nmSMWSCLFt5A1BY1Apk0vX1WLkKkXfmVSMC2BDLvfRi/5jRLXCcVd99n/x0pgjlwrqvIQIIzkwzGVlF24zEzYZ0jGln20XBE9Eq39GHF+y1p2MLkLHwzM/Udz8Klx9IOkgSZMqKrFZbbkYVQ0h2fjVCZLqRjKzCQw/IP+hAh3oCxuJzfk58/JlwBbsjZEnZjEouusQV8K01w6cXCk1AEgDNpvDq62/SNaEE74ZNQ85kImyXlZbccD1z9AAxXVffeNPNXs0lmUDVA4oQhcBk0oqFXWedM3T6tPw9xoOSZCsUIlnX+NuFF9VM+ru+qp6xgJVKOYyvw9jwggIQBqhOAqoolfS6cq65Vng95LAiqfTKq65m1avJ6wXLBGm3+1EKqTvnBey0lRssu/suIy8HFCFyadmV118vqirJnalHRDNlB3w0YEArHERFKtXc1VObzQ3JcbsVn35aRkcZj/z8S+OUu72uQinXuWsPIkhF4cZoyOO77saK6e+6Dj+E0sn2efHkwBuUymqsZYdNrLj/fiQiImRoNjevuvVmI61IFw5VY0VqUn6v5/Irh374vvuE4+xUimHHIAVywRDTjSvYHhMqHnyYCU6KENBOJKpuuNloiILL3aUKcWHHGuTee+cec5RjRnsRxhZrITYAgbG6Z5+zZ35i+B3z3HWUW4JlpQpvutc/YoST4gOIVffep337Iw/lEtld/SlRixPS6uVyzQxXydKSosn3Fp30dyeaDZyhEHXf/Vh5wbni81mGvxA1XTU2aKYNtgLhpERzKCpMaWmdNGqXfkZMN1KNK3znXVi4956gJCICY7WPPGq/9R9PaKCye9pFhBgT6Wit/8JrcnfdpfWCNQ8/av77bVdryAlR2lIPhtwtVYwAkKqqlqm0YGLNSiwwk6bPGHznrcLB95zLZLr28qsMOw0BH+/atrVbQwSUpAzOjz9x2D/O84/cEgC0QQMilPCwAgU2tTTIQFJWuNbS0H/BeQNuuE4L+jPvMOM19zxAH32EwVJSCgiIyH3SpLxzzwtuuzUAuCoqYkRADAQCKQQGAFwpK9qUVKbvxJPLb7nFVVqUgRaMVz/+VOqtf7v8+WDLLslxUnGNys47W7hdJC3kWi+mQYu1Rrljv8yP3Hmf7gpQ91SL0GS4Rhx6aOFfj82cpMgqp89YetedPmKyqbkrMp0AOOMBrxecUjMumCWT0VW4xx5D7r3Xv/VoIEW2RE0jgNWPPV57w81GdT0PlZKUjvuvlG3atiE05+E1TdOZq6ObxYWMRdLDNt/86stbcxhivy1tumeqYWRlafbgANFkcyOMHF187tlAmQs2L1xcecMNQcPXatYRUSpphEKunGDbujY1UjIBRk7Hs1qwVLQ+/8qr83beobUuY8lDD6969y0v+Kgh3NUGEJAQRsDtykQVOFfRBtxt7yGPPSwAyLaR8byDD6oaPjIxf74AxkCzIUUAJhPeA/cpOv3M3MMORgCyJTIELlY98GD4jjt8nmKpCBijWFSM23bggw9qHpdDuuVO2L1m3NaJ2V8a4EZgFqQJlAUgdtqp5Oyzio77G2OZqCpyvvpfrzVcdYXLE+ymlB01Ld200nPIkQX770+y1RfEPrXQSFKpheef76pagZl4bJfBDkjGEsW5Q26/ibtdLf4Z5Y4ePfbHn7jh6SpcB6QQWfTXBbVnnKmFY6i7KR5JauC9+OLyG2/QvB6ybURETUtU1y459WQ17T23K4TB/BYcxpwQiYpFwZXnbLfm9YFAym68gQiWldah/OFHjKIix80igKorrsalS1modK1h9k5Apy2TgorunOwuKlS2zTiXaXP1xZd5GyIQKoJ2wXapBYPMMFobNiSXr0CLo4tR1muHXJPhJth1fPHll4FS0MJklx11RPFhh3Cut0sSbG93UNNqXn+j6dILfYFSJS0iYCASsTiZaRKaw4IbxcUjpr9b++pr0eVLidDr94qKzUJjx3pGj2acAUlS5FRMrbzv0boLLvR7QrZgmLFNIJPNKpUil46MAYKemzv832/WvvKv6G9LCMhmPDRsaM7Y7dyjt9I8LgAimQl917z4ct3Jp7mYB3jXIXRklIpRUVHpTdcxzpz6XOiViG7jgrzq5VfEh59ogULVfWcCZIl0U/45t+RsuWUm0w0QANylJVBashZKEGDVXXfx6joWDKWaqqzB5QMffrjggH0BgEwTdR0AwjM/W3nB+WzOT25/sUJoy2Z2VCSdpmSy9X3GQAAMnVJmW40dZyoa9l1wTsGEPUApAALEquf/mfzPay5/Ecmet5Pk3Aqvyrng4tIDD4KWMErTG/+xPphu+NrHR5zAW0GxA4gdfVRLliCYxDnYKguO20nBBlx7re7k77cgSM+g8rXejh2PJV97xcUzFhqVApdHm784PX+httVohysEInf5wEEXX9gJWjHTqBvIIR2OVN5wY+KBJ/2ePCUESgkIoBQ3XOlFldE5c/In7EGWhZoGREZx0cDzz+tkQ02T6TpybidTq+6cErvnXg/pSjegywJqQo5mLOq77JrAmFGQlVLSC2FdajNisrIqfPNkQ7hVd8w2oRBWuIqN3730ootISWCMwHEtiJQiKbv6UWYaAOpeeCn5+FPc7U00VYu9Jwz7YEbBAfuStEjaqOvSlituvn3FwYeInxZqoVJFsoOxR8bsVNKOJ0gp5xuFz4cuoy1AwwXFwrGyguILL3CoUuQitnxl5XmXu8BDyHrcYY1zFWmkMWNKL76QlCKlkItU5erqG2/SLY246JBCqMjWh1U4INch/yDeDCCzN5dxkYpXB0/9e8E+e5OUwBgROLRJd2to22RZJOXKm29TX83mnhxyFoeIDCEamlbceptMmSgESEmkSLX+rU22lSnwBkTdUJZdN+29RYf8JX7fVN3jVoKDkq196JSm8USk7q67rEQSNQ0yS93+TmzbMQ1M15VUkU8/W3zokdEbbtBsrtxu6Mb74kLFm9PDhuadfPL6d14WnesoACDW3jkFf/0FQ2XdncjIIJWigYMrHnlUeFydcVDdBBR5cuWq2iuvMIDLHG9gyuQhZ53aev4CQGLZ8hUXXGj+598uVz74O0+aY4yrhG1LC1mmiYTwuNFwtWEDJZOaGnT/vd7yMucFkKa95LIr3JEqDJV1nojX7WGElkUuT+Ett+sDypw7IAXLLrtMLZjXyQUJFBAfMAAZazUfqfpGBNGGhxmTqYg9bNSQ2+9qTQxqQ5Bri5M1THsvcv+9Hm8eZbMHUoHfL1/518JotPzeqd5hQ7u6ihWJRr7+umbKFPn+xzroLDSApOyIDZTkgTzzvQ9/mzix6I4poTGju7paMhZv/mBm/VOP0X9nCALDX9xWfNXlikLMTuSdcU6gfECGS1gP6SQ5yQnP1E+fvuqIo9zgBsa770mFiaQaNdx99GGsR2c3KTBc8Y8/o/c+ZG6PNWpY6O9/pUhza3CIAJpffR1/+lEPOoZZdUXEmCS1447wbj6U4gk0dDMSNZ//FzaGwZlUYJqqIOA+/RTOOdg2Gnqyqtp++kWuiBjvsUlABMtSuUHjhL9qfi+YNrqM9PJVseeed6NBncVybJnEA/YNjNueUikUwkomY//3gl4bBrcr81CcqViM77mLe58JkE479MO6rSGBpkdfexN/nstcgTULmZDxdHONGjTQd/zx/vG7uQcOYH4v41ylTbuyKllXk/jq69QHH6d/+VFLk+HPU4y6CZsjF1akShYVBU44wdh5J8+wYcLrRqGBkrK2PrpocfOcH+VXX5qffaUrS/fnq07DRh0tqqBwgzl04JY/ztHcRssLjH2n0EoBgBWPzx83nv+6QATzSdpr3WMiacUaejH3SgMP8wcISEnTTjRCpmdhRlM15ke/b+0YF0HGw0qZAEhAHAR354Boy54jaduJBoJM5AmZJjw562EEACTJRJMEGx0NRqH7cgm66tyFMh1TZpwyGX1M8+V2BHvIIJW2rKaen7gkXEGmeZyIQafENsUT0oqCx28PKCGuIQCmTWyOUrgebFMIL9ddoOlr32gA5JzMpJVsIsMPRSXK63Ky1rCpSVVXM7AF93CXD4S2LlcDRFQqIaNlL75YMPGwlpoU7EsL7RjI5ddcF73lNiNUotb5REbWm5Mi05G2iytk/3Zt395uIdbMcOpwcVrvuQJr/cYOn85OGev8wwiIvVrDDqVNnZ6iTm6GmUZLZfJPXAKZ5pQLUI9aNTvPQgpTacecE0MyBPIMbdp5mkcX3oiM1GpHHzns5RcBHGaDesfWdaHQJAF5+PsfVu61v5GSpOug+sejbEKSXSKlaH0nSbIsbNCLqyGCLW2XKv/vf/077dACNan7epyeOIVEQCBNs+qaa3k4QsH8TN5qv2wyQtSX01B7b+zQyZQ0k7Xa30/17riDwxG3Baf7hOUgImS87qmnrHenuXMHkZLQLtW4X/qlD48KRjKVDvnLL7mY9anRFK2+IDJW/+33K664xAWUaFwJ0D9csF82mEIDs0Dm33ZHYNgwyjTi76MrZ2Po6PyF6blzma73z2Lrlw0L5YmIi8D48VrQn/mXDaHQ/dIvG7uINZyGfv3ul9/RVGMf0w79FrpfNilh/UvQL5uS/D9/6wKEl8KUtAAAAABJRU5ErkJggg==';   // bản chia sẻ: ảnh logo (data URI) làm nút mở tool ở góc trên bên trái; để trống thì dùng nút chữ ở góc dưới bên phải
  const RAW_FETCH = typeof window.fetch === 'function' ? window.fetch.bind(window) : null;   // hàm fetch gốc, lấy trước khi tool gắn lớp bắt dữ liệu
  window.__b8sn = (window.__b8sn || 0) + 1;   // đếm số bản script cùng chạy trong trang (phát hiện cài trùng)
  const state = {
    office: '', province: '', officeList: [], recent: [], listTried: false,
    lastRefresh: null, busy: false, view: 'dash',
    ov: null,                  // số liệu trang tổng quan, chỉ dùng để đối chiếu
    emp: null,                 // {headers, rows, at}  tab Tiến độ nhân viên
    orders: null,              // {headers, rows, wbCol, at}  tab Chi tiết phát kiện
    sort: { key: 'rateTrip', dir: -1 },
    tiktok: 'TTVN',            // mã đầu vận đơn của TikTok Shop (cách nhau dấu phẩy)
    statusMap: {},             // trạng thái -> ok | fail | unscanned | other
    includePII: false,         // luôn tắt: không giữ SĐT/địa chỉ người nhận
    hist: new Map(), histReady: false, histErr: '', pending: null, lastSaved: null,
    hv: { office: '', mode: 'trend', preset: 'd30', from: '', to: '', group: 'day', cmpDate: '', ssort: { key: 'rateTrip', dir: -1 } },
    anm: { sel: null, base: '', cmp: 'slot2', slot: '', lastAuto: 0, authErr: '', rankKey: 'pod', trendKey: 'pod', hide: new Set(), open: '', loading: null, results: [], detail: '' },
    cod: { by: {}, meta: {}, err: {}, warn: {}, running: false, at: 0, open: false, copied: 0, zero: '', zcopied: 0 },   // COD ship nộp tiền (chỉ giữ trong bộ nhớ, không lưu)
    ai: { key: '', model: 'claude-haiku-5-5', keyOk: false, on: false, prompt: '', res: null, running: false, err: '', test: { busy: false } },   // dự báo POD bằng AI
    wx: { data: null, at: 0, err: '', loading: false, key: '' },   // thời tiết (lượng mưa theo giờ); on/lat/lon nạp từ cấu hình
    sign: null,                // chẩn đoán trường giờ ký nhận: { keys, okN, signN }
    snap: { list: [], idx: new Map(), ready: false, running: false, tries: {}, last: {}, err: '', timer: 0 },   // bản chụp theo mốc giờ của tab ANM
    fast: true, apiHdr: (window.__b8shdr = window.__b8shdr || {}), sites: {}, apiNote: '', ownerSiteId: null,   // tải nhanh bằng API (chỉ ANM): tiêu đề xác thực lấy từ chính trang
    th: {},                    // ngưỡng màu/cảnh báo do người dùng chỉnh (ghi đè mặc định)
    dm: null,                  // quy tắc tính ship đi làm (nạp từ cấu hình, mặc định DM_RULES_DEFAULT)
    calc: null, calcKey: '',
    endpoints: new Map()       // chỉ lưu tên khóa API để chẩn đoán, KHÔNG lưu giá trị
  };


  /* ------------------------------------------------------------------ *
   * KHÓA TAB ANM (bản chia sẻ). Mã của tab ANM được mã hóa (AES-256-GCM, khóa sinh từ mật khẩu bằng PBKDF2) và đặt trên GitHub.
   * Tổng quan và Cài đặt luôn dùng được; bấm vào ANM thì hỏi mật khẩu. File cấu hình status.json do người quản lý sửa:
   *   enabled (bật/tắt), minVersion, pwVer (đổi số này = mọi người phải nhập lại), kdf (salt, iter), payloadVer, graceDays, msg.
   * Chỉ đọc 2 file công khai trên GitHub; KHÔNG gửi bất kỳ dữ liệu nào của người dùng đi đâu.
   * ------------------------------------------------------------------ */
  const LK = { user: 'Erick1939', repo: 'tool-best', branch: 'main', ver: '3.15.1', updateUrl: 'https://raw.githubusercontent.com/Erick1939/tool-best/main/800best-share.user.js', graceDefault: 3, SKEW: 600000, RECHECK: 1800000, MAXFAIL: 5 };
  const lkS = { st: 'init', msg: '', busy: false, status: null, fails: 0, wait: 0, diag: '', checkedAt: 0 };   // st: init | checking | locked | blocked | error | open
  const lkNoop = () => undefined, lkResolved = () => Promise.resolve();
  const lkB64d = s => Uint8Array.from(atob(String(s)), c => c.charCodeAt(0));
  const lkB64e = u => { let s = ''; u.forEach(b => { s += String.fromCharCode(b); }); return btoa(s); };
  const lkVerCmp = (a, b) => { const x = String(a).split('.').map(Number), y = String(b).split('.').map(Number); for (let i = 0; i < 3; i++) { const d = (x[i] || 0) - (y[i] || 0); if (d) return d < 0 ? -1 : 1; } return 0; };
  const lkUrl = n => `https://raw.githubusercontent.com/${LK.user}/${LK.repo}/${LK.branch}/${n}?t=${Math.floor(Date.now() / 60000)}`;
  function lkStore(o) { try { if (o === undefined) return JSON.parse(localStorage.getItem('b8s_lk') || 'null') || {}; localStorage.setItem('b8s_lk', JSON.stringify(o)); } catch { /* bỏ qua */ } return {}; }
  const lkPaySave = t => { try { localStorage.setItem('b8s_lkpay', t); } catch { /* bỏ qua */ } };
  const lkPayLoad = () => { try { return localStorage.getItem('b8s_lkpay') || ''; } catch { return ''; } };
  async function lkGet(name) {
    const f = typeof RAW_FETCH === 'function' ? RAW_FETCH : (typeof fetch === 'function' ? fetch.bind(window) : null);
    if (!f) throw new Error('Trình duyệt không hỗ trợ tải cấu hình');
    let r;
    try { r = await f(lkUrl(name), { cache: 'no-store', credentials: 'omit' }); }
    catch (e) { throw Object.assign(new Error('Không kết nối được tới GitHub (mạng hoặc trang đang chặn kết nối ra ngoài)'), { net: true, detail: String(e && e.message || e) }); }
    if (!r.ok) throw Object.assign(new Error(r.status === 404 ? 'Không thấy file cấu hình trên GitHub (HTTP 404)' : `GitHub trả lỗi HTTP ${r.status}`), { net: r.status >= 500, http: r.status });
    return r.text();
  }
  function lkParseStatus(txt) {
    let o; try { o = JSON.parse(txt); } catch { throw new Error('File cấu hình trên GitHub bị hỏng (không đọc được)'); }
    if (!o || typeof o !== 'object' || !o.kdf || typeof o.kdf.salt !== 'string' || !(o.kdf.iter >= 1000) || !Number.isFinite(Number(o.pwVer))) throw new Error('File cấu hình trên GitHub sai định dạng');
    return o;
  }
  // Quyết định theo cấu hình: trả {ok:true} hoặc {ok:false, why}
  function lkJudge(st) {
    if (st.enabled === false) return { ok: false, why: String(st.msg || 'Tab ANM đang tạm khóa bởi người quản lý.') };
    if (st.minVersion && lkVerCmp(LK.ver, st.minVersion) < 0) return { ok: false, why: `Bản tool này (${LK.ver}) đã cũ, cần bản ${st.minVersion} trở lên.${LK.updateUrl ? ' Hãy cập nhật tool trong Tampermonkey.' : ''}` };
    return { ok: true };
  }
  async function lkRefreshStatus() {
    const txt = await lkGet('status.json'), st = lkParseStatus(txt), now = Date.now(), sv = lkStore();
    lkS.status = st; lkS.checkedAt = now;
    lkStore(Object.assign({}, sv, { status: st, lastOk: now, maxSeen: now }));   // vừa kết nối được GitHub: lấy giờ máy hiện tại làm mốc mới
    return st;
  }
  async function lkImportKey(raw) { return crypto.subtle.importKey('raw', raw, 'AES-GCM', true, ['decrypt']); }
  async function lkDerive(pw, st) {
    const base = await crypto.subtle.importKey('raw', new TextEncoder().encode(pw.normalize('NFC')), 'PBKDF2', false, ['deriveBits']);
    return new Uint8Array(await crypto.subtle.deriveBits({ name: 'PBKDF2', salt: lkB64d(st.kdf.salt), iterations: Number(st.kdf.iter), hash: 'SHA-256' }, base, 256));
  }
  // Giải mã gói mã hóa -> chuỗi mã nguồn. Ném lỗi {bad:true} nếu khóa sai.
  async function lkDecrypt(rawKey, payText) {
    let p; try { p = JSON.parse(payText); } catch { throw new Error('Gói mã hóa trên GitHub bị hỏng'); }
    if (!p || !p.iv || !p.ct) throw new Error('Gói mã hóa trên GitHub sai định dạng');
    let buf;
    try { buf = await crypto.subtle.decrypt({ name: 'AES-GCM', iv: lkB64d(p.iv), additionalData: new TextEncoder().encode('anm-v1') }, await lkImportKey(rawKey), lkB64d(p.ct)); }
    catch { throw Object.assign(new Error('Sai mật khẩu'), { bad: true }); }
    if (typeof DecompressionStream !== 'function') throw new Error('Trình duyệt không hỗ trợ giải nén (cần Chrome/Edge bản mới)');
    const ds = new Blob([buf]).stream().pipeThrough(new DecompressionStream('gzip'));
    return new Response(ds).text();
  }
  function lkMark() { const a = $('#b8s-app'); if (a) a.classList.toggle('lk', lkS.st !== 'open'); }
  function lkActivate(src) {
    if (!/^\/\*B8ANM/.test(src)) throw new Error('Gói mã hóa không đúng loại');
    let o;
    try { o = eval('(' + src + ')')(); }   // eslint-disable-line no-eval -- chạy trong phạm vi của tool để dùng chung các hàm
    catch (e) { throw Object.assign(new Error(`Không chạy được mã ANM (${e && e.message || e}). Có thể trang chặn chạy mã động (CSP).`), { csp: /eval|unsafe|Content Security/i.test(String(e && e.message)) }); }
    lkBind(o); lkS.st = 'open'; lkS.msg = ''; lkMark();
    try { loadAi(); loadSnapLog(); loadSnaps().then(() => { renderUI(); snapBadge(); }); } catch { /* bỏ qua */ }
    return true;
  }
  function lkRelock(st, why) {
    const was = lkS.st === 'open';
    lkUnbind(); lkS.st = st; lkS.msg = why || ''; lkMark();
    if (was) { try { renderUI(); } catch { /* bỏ qua */ } }
  }
  // Vào ứng dụng: nếu đã nhập đúng trước đó và còn trong hạn thì mở ngay, rồi kiểm tra lại với GitHub ở nền
  async function lkBoot() {
    lkMark();
    const sv = lkStore(), now = Date.now();
    const graceMs = (Number(sv.status && sv.status.graceDays) || LK.graceDefault) * 86400000;
    const cached = sv.status && sv.key && sv.status.pwVer === sv.pwVer && lkPayLoad();
    const clockOk = now >= (sv.maxSeen || 0) - LK.SKEW;
    if (cached && clockOk && now - (sv.lastOk || 0) <= graceMs && lkJudge(sv.status).ok) {
      try { lkActivate(await lkDecrypt(lkB64d(sv.key), lkPayLoad())); renderUI(); }
      catch (e) { lkS.st = 'locked'; lkS.msg = e.bad ? '' : e.message; lkMark(); }
    } else {
      lkS.st = 'locked';
      if (sv.key && sv.status && sv.status.pwVer === sv.pwVer) lkS.msg = !clockOk ? 'Đồng hồ máy bị lùi so với lần dùng trước. Hãy kết nối mạng để kiểm tra lại.' : `Đã quá ${Math.round(graceMs / 86400000)} ngày chưa kết nối được GitHub để kiểm tra. Hãy kết nối mạng, tool sẽ tự mở lại.`;
    }
    lkStore(Object.assign({}, lkStore(), { maxSeen: Math.max(now, (lkStore().maxSeen || 0)) }));
    lkRecheck(); setInterval(lkRecheck, LK.RECHECK);
    const pn = $('#b8s-anm'); if (pn) pn.addEventListener('keydown', e => { if (e.key === 'Enter' && e.target && e.target.id === 'b8s-lk-pw') lkGo(); });
  }
  // Kiểm tra lại cấu hình trên GitHub: khóa lại nếu bị tắt, đổi mật khẩu, quá hạn hoặc cần cập nhật
  async function lkRecheck() {
    if (lkS.busy) return;
    if (lkS.st === 'init') return;
    let st;
    try { st = await lkRefreshStatus(); }
    catch (e) {
      const sv = lkStore(), now = Date.now(), graceMs = (Number(sv.status && sv.status.graceDays) || LK.graceDefault) * 86400000;
      if (lkS.st === 'open' && (!(now - (sv.lastOk || 0) <= graceMs) || now < (sv.maxSeen || 0) - LK.SKEW)) lkRelock('locked', 'Đã quá hạn dùng khi không kết nối được GitHub. Hãy kết nối mạng rồi nhập lại mật khẩu.');
      else if (lkS.st === 'open') lkStore(Object.assign({}, sv, { maxSeen: Math.max(now, sv.maxSeen || 0) }));
      lkS.diag = e.detail ? `${e.message}: ${e.detail}` : e.message;
      if (state.view === 'anm') renderAnm();
      return;
    }
    const sv = lkStore(), j = lkJudge(st);
    if (sv.key && Number(sv.pwVer) !== Number(st.pwVer)) { lkStore(Object.assign({}, sv, { key: '', pwVer: null })); lkRelock('locked', 'Mật khẩu đã được đổi. Hãy nhập mật khẩu mới.'); }
    else if (!j.ok) lkRelock('blocked', j.why);
    else if (lkS.st === 'blocked' || lkS.st === 'locked') {
      // đã nhập đúng từ trước và cấu hình vẫn hợp lệ (vừa hết chặn / vừa có mạng lại): mở lại không cần nhập
      if (sv.key && Number(sv.pwVer) === Number(st.pwVer) && lkPayLoad()) { try { lkActivate(await lkDecrypt(lkB64d(sv.key), lkPayLoad())); } catch { lkS.st = 'locked'; } } else if (lkS.st === 'blocked') lkS.st = 'locked';
    }
    const sv2 = lkStore();
    if (lkS.st === 'open' && st.payloadVer && st.payloadVer !== sv2.pv && sv2.key) {
      // có bản mã mới: tải về, kiểm tra giải mã được rồi lưu để dùng từ lần mở sau
      try { const t = await lkGet('anm.dat'); await lkDecrypt(lkB64d(sv2.key), t); lkPaySave(t); lkStore(Object.assign({}, lkStore(), { pv: st.payloadVer })); log('Có bản mới của tab ANM, sẽ áp dụng khi bạn tải lại trang.'); }
      catch { /* giữ bản cũ */ }
    }
    if (state.view === 'anm' || lkS.st !== 'open') { try { renderUI(); } catch { /* bỏ qua */ } }
  }
  async function lkGo() {
    if (lkS.busy) return;
    const inp = $('#b8s-lk-pw'), pw = inp ? inp.value : '';
    if (Date.now() < lkS.wait) { lkS.msg = `Nhập sai nhiều lần. Thử lại sau ${Math.ceil((lkS.wait - Date.now()) / 1000)} giây.`; renderAnm(); return; }
    if (!pw) { lkS.msg = 'Hãy nhập mật khẩu.'; renderAnm(); return; }
    lkS.busy = true; lkS.msg = ''; renderAnm();
    try {
      if (!(window.crypto && crypto.subtle)) throw new Error('Trình duyệt không hỗ trợ mã hóa (cần mở trang bằng https)');
      const st = await lkRefreshStatus(), j = lkJudge(st);
      if (!j.ok) { lkS.st = 'blocked'; lkS.msg = j.why; return; }
      const raw = await lkDerive(pw, st), sv = lkStore();
      let pay = sv.pv && sv.pv === st.payloadVer ? lkPayLoad() : '';
      let src = '';
      if (pay) { try { src = await lkDecrypt(raw, pay); } catch (e) { if (!e.bad) throw e; pay = ''; } }
      if (!src) { pay = await lkGet('anm.dat'); src = await lkDecrypt(raw, pay); }
      lkPaySave(pay);
      lkStore(Object.assign({}, lkStore(), { key: lkB64e(raw), pwVer: Number(st.pwVer), pv: st.payloadVer || '', lastOk: Date.now() }));
      lkS.fails = 0; lkActivate(src); lkS.msg = '';
      if (inp) inp.value = '';
      log('Đã mở khóa tab ANM.');
    } catch (e) {
      if (e.bad) { lkS.fails++; lkS.msg = 'Sai mật khẩu.'; if (lkS.fails >= LK.MAXFAIL) { lkS.wait = Date.now() + 30000 * (lkS.fails - LK.MAXFAIL + 1); lkS.msg += ` Đã sai ${lkS.fails} lần, tạm khóa nhập ${30 * (lkS.fails - LK.MAXFAIL + 1)} giây.`; } }
      else { lkS.msg = e.message; lkS.diag = e.detail ? `${e.message}: ${e.detail}` : e.message; }
    } finally { lkS.busy = false; try { renderUI(); } catch { /* bỏ qua */ } }
  }
  function lkDiagText() {
    let ev = 'không'; try { new Function('return 1')(); ev = 'có'; } catch { /* bị chặn */ }
    return `Chạy mã động: ${ev} · mã hóa (crypto.subtle): ${window.crypto && crypto.subtle ? 'có' : 'KHÔNG'} · giải nén: ${typeof DecompressionStream === 'function' ? 'có' : 'KHÔNG'} · kho: ${LK.user}/${LK.repo} · bản tool ${LK.ver}${lkS.diag ? ' · lỗi gần nhất: ' + lkS.diag : ''}`;
  }
  function lkRenderAnm() {
    const box = $('#b8s-anm'); if (!box) return;
    const wait = Math.max(0, Math.ceil((lkS.wait - Date.now()) / 1000));
    const blocked = lkS.st === 'blocked';
    const html = `<div class="b8s-panel" style="max-width:520px;margin:40px auto;text-align:center;padding:26px"><h3 style="margin:0 0 6px">🔒 Tab ANM đã khóa</h3>`
      + (blocked ? `<div class="al a" style="margin:10px 0">${esc(lkS.msg)}</div>` : `<div class="muted" style="font-size:14px;margin-bottom:12px">Nhập mật khẩu để mở tab ANM. Tab Tổng quan và Cài đặt vẫn dùng bình thường.</div>`
        + `<div style="display:flex;gap:8px;justify-content:center"><input type="password" id="b8s-lk-pw" autocomplete="off" placeholder="Mật khẩu" style="width:230px;padding:7px 10px" ${lkS.busy ? 'disabled' : ''}><span class="btn" data-lk="go">${lkS.busy ? '⟳ Đang kiểm tra…' : 'Mở khóa'}</span></div>`
        + (lkS.msg ? `<div class="al ${/Sai|sai|Không|không|lỗi|quá hạn|đổi/.test(lkS.msg) ? 'r' : 'a'}" style="margin:12px 0 0">${esc(lkS.msg)}</div>` : '') + (wait ? `<div class="muted">Chờ ${wait} giây…</div>` : ''))
      + `<div style="margin-top:14px"><small class="muted" data-lk="diag" style="cursor:pointer;text-decoration:underline">Chẩn đoán</small></div><div id="b8s-lk-diag" class="muted" style="font-size:11.5px;margin-top:6px;word-break:break-word"></div></div>`;
    if (box.dataset.lk !== html) { const keep = $('#b8s-lk-pw') ? $('#b8s-lk-pw').value : ''; box.dataset.lk = html; box.innerHTML = html; const n = $('#b8s-lk-pw'); if (n && keep) n.value = keep; }
  }
  function lkClick(e) {
    const el = e.target.closest('[data-lk]'); if (!el) return;
    if (el.dataset.lk === 'go') lkGo();
    else if (el.dataset.lk === 'diag') { const d = $('#b8s-lk-diag'); if (d) d.textContent = lkDiagText(); }
  }
  function lkEnter() {
    if (lkS.st === 'open') return;
    renderAnm();
    if (!lkS.busy && Date.now() - lkS.checkedAt > 60000) lkRecheck();
  }

  // ----- Tab ANM bị khóa: các hàm của ANM là stub cho tới khi nhập đúng mật khẩu (xem lock_loader) -----
  let loadSnaps = lkResolved;
  let loadSnapLog = lkNoop;
  let snapTick = lkNoop;
  let snapBadge = lkNoop;
  let loadWeather = lkNoop;
  let renderWxInfo = lkNoop;
  let onEnterAnm = lkEnter;
  let loadAi = lkNoop;
  let saveAi = lkNoop;
  let renderAiInfo = lkNoop;
  let aiSubmitKey = lkNoop;
  let aiClearKey = lkNoop;
  let fitAnm = lkNoop;
  let renderAnm = lkRenderAnm;
  let onAnmClick = lkClick;
  let onAnmChange = lkNoop;
  let addAnmSheets = lkNoop;
  function lkBind(o) { loadSnaps = o.loadSnaps; loadSnapLog = o.loadSnapLog; snapTick = o.snapTick; snapBadge = o.snapBadge; loadWeather = o.loadWeather; renderWxInfo = o.renderWxInfo; onEnterAnm = o.onEnterAnm; loadAi = o.loadAi; saveAi = o.saveAi; renderAiInfo = o.renderAiInfo; aiSubmitKey = o.aiSubmitKey; aiClearKey = o.aiClearKey; fitAnm = o.fitAnm; renderAnm = o.renderAnm; onAnmClick = o.onAnmClick; onAnmChange = o.onAnmChange; addAnmSheets = o.addAnmSheets; }
  function lkUnbind() { loadSnaps = lkResolved; loadSnapLog = lkNoop; snapTick = lkNoop; snapBadge = lkNoop; loadWeather = lkNoop; renderWxInfo = lkNoop; onEnterAnm = lkEnter; loadAi = lkNoop; saveAi = lkNoop; renderAiInfo = lkNoop; aiSubmitKey = lkNoop; aiClearKey = lkNoop; fitAnm = lkNoop; renderAnm = lkRenderAnm; onAnmClick = lkClick; onAnmChange = lkNoop; addAnmSheets = lkNoop; }

  /* ------------------------------------------------------------------ *
   * Hàm thuần (không đụng DOM) – có test riêng
   * ------------------------------------------------------------------ */
  // <pure>
  const esc = v => String(v ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

  function parseNum(v) {
    const s = String(v ?? '').trim();
    if (!s) return NaN;
    const m = s.match(/-?\d[\d.,]*/);
    if (!m) return NaN;
    let t = m[0];
    if (/^-?\d{1,3}([.,]\d{3})+$/.test(t)) t = t.replace(/[.,]/g, '');
    else t = t.replace(',', '.');
    return parseFloat(t);
  }

  const fmt = n => (typeof n === 'number' && Number.isFinite(n)) ? n.toLocaleString('vi-VN') : String(n ?? '');
  const round1 = v => Math.round(v * 10) / 10;
  const ratio = (a, b) => (b > 0 && Number.isFinite(a)) ? a / b * 100 : null;

  function fmtVal(v, kind) {
    if (v === null || v === undefined || !Number.isFinite(v)) return '—';
    if (kind === 'pct') return round1(v).toLocaleString('vi-VN') + '%';
    if (kind === 'dec') return round1(v).toLocaleString('vi-VN');
    return Math.round(v).toLocaleString('vi-VN');
  }

  function uniqueHeaders(list, n) {
    const seen = new Map();
    return Array.from({ length: n }, (_, i) => {
      const h = String(list[i] || `Cột ${i + 1}`).normalize('NFC');
      const c = (seen.get(h) || 0) + 1;
      seen.set(h, c);
      return c > 1 ? `${h} (${c})` : h;
    });
  }

  function toCell(v) {
    if (typeof v === 'number') return v;
    const s = String(v ?? '').trim();
    if (!/^-?\d[\d.,]*$/.test(s)) return s;
    if (/^0\d/.test(s) || s.replace(/\D/g, '').length > 9) return s;
    const n = parseNum(s);
    return Number.isFinite(n) ? n : s;
  }

  const pad = n => String(n).padStart(2, '0');
  const localStamp = d => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}_${pad(d.getHours())}${pad(d.getMinutes())}`;
  const safeName = s => String(s || '').replace(/[\\/:*?"<>|\s]+/g, '_').slice(0, 40);

  // ---- Vận đơn / khách ----
  const WAYBILL_RE = /^(?:[A-Za-z]{2,6}[A-Za-z0-9]{6,}|\d{10,})$/;
  const isWaybill = v => { const t = String(v ?? '').trim(); return WAYBILL_RE.test(t) && /\d/.test(t); };

  function waybillCol(headers, rows) {
    const byHeader = headers.findIndex(h => /vận đơn|waybill|运单/i.test(h));
    if (byHeader >= 0) return byHeader;
    let best = -1, bestScore = 0;
    headers.forEach((_, c) => {
      const vals = rows.slice(0, 30).map(r => r[c]);
      const score = vals.filter(isWaybill).length / Math.max(1, vals.length);
      if (score > bestScore) { bestScore = score; best = c; }
    });
    return bestScore >= 0.6 ? best : -1;
  }

  const colBy = (headers, re) => headers.findIndex(h => re.test(h));
  const parseList = text => String(text || '').split(/[,;\s]+/).map(x => x.trim().toUpperCase()).filter(Boolean);
  const isTikTok = (wb, prefixes) => { const w = String(wb ?? '').trim().toUpperCase(); return !!w && prefixes.some(p => w.startsWith(p)); };

  function sheetName(name, used) {
    const base = String(name || 'Sheet').replace(/[\[\]:*?\/\\]/g, ' ').trim().slice(0, 31) || 'Sheet';
    let out = base, i = 2;
    while (used.has(out.toLowerCase())) { const suf = ` ${i++}`; out = base.slice(0, 31 - suf.length) + suf; }
    used.add(out.toLowerCase());
    return out;
  }

  // ---- Ngày / nhân viên ----
  const DATE_RE = /(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{4})|(\d{4})[\/.\-](\d{1,2})[\/.\-](\d{1,2})/;
  function parseDate(s) {
    const m = String(s ?? '').match(DATE_RE);
    if (!m) return null;
    return m[1] ? { d: +m[1], m: +m[2], y: +m[3] } : { d: +m[6], m: +m[5], y: +m[4] };
  }
  const isToday = (s, now) => {
    const p = parseDate(s);
    return !!p && p.d === now.getDate() && p.m === now.getMonth() + 1 && p.y === now.getFullYear();
  };

  // Số ngày từ ngày trong ô đến hôm nay (dương = ngày trong quá khứ); null nếu không đọc được ngày
  const dayNum = p => Math.round(Date.UTC(p.y, p.m - 1, p.d) / 86400000);
  function daysBefore(s, now) {
    const p = parseDate(s);
    return p ? dayNum({ y: now.getFullYear(), m: now.getMonth() + 1, d: now.getDate() }) - dayNum(p) : null;
  }

  // Tên không dấu, bỏ mã số/ngoặc để so khớp giữa hai tab
  const normName = s => String(s ?? '').replace(/\(.*?\)/g, ' ').replace(/[\d.]+/g, ' ')
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/đ/g, 'd').replace(/Đ/g, 'D')
    .toLowerCase().replace(/[^a-z ]/g, ' ').replace(/\s+/g, ' ').trim();

  function shipKeys(s) {
    const raw = String(s ?? '').trim();
    if (!raw) return { code: '', name: '' };
    const m = raw.match(/\d{4,}(?:\.\d+)?/);
    return { code: m ? m[0] : '', name: normName(raw) };
  }

  // ---- Phân loại trạng thái đơn ----
  // ok = giao thành công | unscanned = nhập kho chưa quét giao | other = còn lại (đã quét giao: đang giao, thất bại, hoàn…)
  function defaultCat(status) {
    const t = String(status ?? '').normalize('NFC').toLowerCase();
    if (/đã ký nhận|giao hàng thành công/.test(t)) return 'ok';
    if (/nhập kho|chưa quét|đã đến|chờ phát|chưa phát/.test(t)) return 'unscanned';
    return 'other';
  }
  function classify(status, map) {
    const s = String(status ?? '').trim();
    const v = map && map[s];
    return (v === 'ok' || v === 'unscanned' || v === 'other') ? v : defaultCat(s);
  }

  // ---- Bưu cục / tỉnh ----
  const plain = s => String(s ?? '').normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/đ/g, 'd').replace(/Đ/g, 'D')
    .toLowerCase().replace(/\s+/g, ' ').trim();
  const titleCase = s => String(s ?? '').toLocaleLowerCase('vi').replace(/(^|\s)(\S)/g, (_, a, b) => a + b.toLocaleUpperCase('vi'));

  function parseProvince(text) {
    const m = String(text ?? '').match(/Khu vực trực thuộc\s*[:：]\s*([^\n]+)/i);
    return m ? m[1].split(/\s{2,}|Thuộc khu vực|所属/)[0].trim().slice(0, 40) : '';
  }
  function parseOfficeText(text) {
    const m = String(text ?? '').match(/(?<![\p{L}])(?<!(?:Mã|Tên|Loại|Cấp)\s)Bưu cục\s*[:：]\s*([^\n]+)/iu);
    return m ? m[1].split(/\s{2,}|Thuộc khu vực|Bao gồm|所属/)[0].trim().slice(0, 60) : '';
  }

  // Chọn một dòng trong danh sách thả xuống: khớp đúng tên trước, rồi khớp phần đầu, rồi khớp duy nhất
  function pickOption(texts, name) {
    const t = plain(name), p = texts.map(plain);
    if (!t) return { index: -1, ambiguous: [] };
    const exact = p.findIndex(x => x === t);
    if (exact >= 0) return { index: exact, ambiguous: [] };
    const hits = p.map((x, k) => x.includes(t) ? k : -1).filter(k => k >= 0);
    if (hits.length === 1) return { index: hits[0], ambiguous: [] };
    const strict = hits.filter(k => p[k].startsWith(t) && !/^\s*[a-z0-9]/.test(p[k].slice(t.length)));
    if (strict.length === 1) return { index: strict[0], ambiguous: [] };
    return { index: -1, ambiguous: hits.map(k => texts[k]) };
  }

  function statusCounts(rows, stCol) {
    const m = new Map();
    for (const r of rows) { const s = String(r[stCol] ?? '').trim() || '(trống)'; m.set(s, (m.get(s) || 0) + 1); }
    return [...m].sort((a, b) => b[1] - a[1]);
  }

  // Shipper đi làm hôm nay: dòng có ngày hôm nay ở cột ngày (cột thứ 4); để trắng = không đi làm
  function workersOf(emp, now) {
    if (!emp || !emp.rows.length) return null;
    const H = emp.headers.map(h => String(h ?? '').normalize('NFC')), R = emp.rows;
    const find = l => { const i = H.findIndex(h => h === l); return i >= 0 ? i : H.findIndex(h => h.includes(l)); };
    const needCol = find('Đơn cần phát'), doneCol = find('Đã quét phát');
    let nameCol = find('Nhân viên giao nhận'); if (nameCol < 0) nameCol = 0;
    const dateCnt = H.map((_, c) => R.filter(r => parseDate(r[c])).length);
    let workCol = dateCnt[3] > 0 ? 3 : dateCnt.findIndex(n => n > 0);
    let list, note = '';
    if (workCol >= 0) {
      list = R.filter(r => isToday(r[workCol], now));
      if (!list.length) {
        list = R.filter(r => String(r[workCol] ?? '').trim() !== '');
        note = `Không dòng nào có ngày hôm nay ở cột “${H[workCol]}”; tạm tính các dòng có ngày.`;
      }
    } else {
      workCol = Math.min(3, H.length - 1);
      list = R.filter(r => String(r[workCol] ?? '').trim() !== '');
      note = 'Không thấy cột ngày; tạm tính các dòng có dữ liệu ở cột thứ 4.';
    }
    return { list, nameCol, needCol, doneCol, workCol, deptCol: find('Chức vụ'), workHeader: H[workCol] || `Cột ${workCol + 1}`, note };
  }


  // ---- Giờ giao thành công: trường thời gian của đơn (tên trường chưa được xác nhận nên thử lần lượt) ----
  const SIGN_KEYS = ['signTime', 'signingTime', 'signDate', 'signedTime', 'signOffTime', 'finishTime', 'finishDeliveryTime', 'deliveryFinishTime', 'deliveredTime', 'deliverTime', 'deliveryTime', 'latestScanTime', 'lastScanTime', 'scanTime'];
  const msOf = v => { const n = (typeof v === 'string' && /^\d{10,13}$/.test(v.trim())) ? Number(v) : v; return (typeof n === 'number' && Number.isFinite(n) && n >= 1e11 && n < 4e12) ? n : null; };
  function signMsOf(it) { for (const k of SIGN_KEYS) { const m = msOf(it && it[k]); if (m !== null) return m; } return null; }
  // Tên các trường trông như mốc thời gian (chỉ tên và số đơn có giá trị, không lấy giá trị) – để chẩn đoán
  function timeKeysOf(items) {
    const cnt = new Map();
    for (const it of (items || []).slice(0, 300)) {
      if (!it || typeof it !== 'object') continue;
      for (const [k, v] of Object.entries(it)) if (msOf(v) !== null || (/time|date/i.test(k) && v !== null && v !== undefined && v !== '')) cnt.set(k, (cnt.get(k) || 0) + 1);
    }
    return [...cnt].sort((a, b) => b[1] - a[1]).slice(0, 40).map(([k, n]) => `${k}(${n})`);
  }
  // Phút trong ngày (0..1439) của chuỗi “dd/mm/yyyy hh:mm:ss”; null nếu không đọc được
  function minOfDay(s) { const m = String(s ?? '').match(/(?:^|\s|T)(\d{1,2}):(\d{2})/); return m && +m[1] < 24 && +m[2] < 60 ? +m[1] * 60 + +m[2] : null; }
  const hmOf = m => Number.isFinite(m) ? `${pad(Math.floor(m / 60))}:${pad(Math.round(m) % 60)}` : '—';

  // ---- Quy tắc tính ship đi làm ----
  // Chỉ tính ship có chức vụ trong danh sách, quét đơn đầu tiên không muộn hơn mốc cutoff; cuối ngày (từ mốc cuối cùng) loại ship giao thành công ít hơn minOk đơn
  const DM_RULES_DEFAULT = {
    on: true,
    depts: ['Nhân viên giao nhận - Chính thức', 'Nhân viên giao nhận - Chính thức (CARGO)', 'Nhân viên giao nhận - Thuê ngoài dài hạn', 'Nhân viên giao nhận - Thuê ngoài ngắn hạn'],
    cutoff: '14:00', minOk: 5
  };
  const normDept = s => plain(String(s ?? '').replace(/\uFF08/g, '(').replace(/\uFF09/g, ')')).replace(/\s*-\s*/g, ' - ').replace(/\s*\(\s*/g, ' (').replace(/\s*\)/g, ')').replace(/\s+/g, ' ').trim();
  const hmToMin = t => { const m = /^(\d{1,2}):(\d{2})$/.exec(String(t ?? '').trim()); return m && +m[1] < 24 && +m[2] < 60 ? +m[1] * 60 + +m[2] : NaN; };
  // Cuối ngày = từ mốc tự lưu cuối cùng trở đi (23:30)
  const dmFinal = now => now.getHours() * 60 + now.getMinutes() >= slotMin(SNAP_SLOTS[SNAP_SLOTS.length - 1]);
  function dmRulesOf(r) {
    const o = Object.assign({}, DM_RULES_DEFAULT, r && typeof r === 'object' ? r : {});
    o.depts = Array.isArray(o.depts) ? o.depts.map(x => String(x).trim()).filter(Boolean) : DM_RULES_DEFAULT.depts.slice();
    o.on = o.on !== false; o.minOk = Number.isFinite(+o.minOk) ? +o.minOk : 5;
    if (!Number.isFinite(hmToMin(o.cutoff))) o.cutoff = '14:00';
    return o;
  }

  // ---- Dữ liệu API -> bảng giống bảng trên trang (để dùng lại computeDashboard) ----
  const STATUS_TXT = { 2: 'Đã nhập kho nhưng chưa quét phát', 3: 'Đang phát hàng', 4: 'Giao hàng thành công', 5: 'Giao hàng thất bại', 22: 'Chờ duyệt', 23: 'Chưa in bill hoàn', 36: 'Từ chối hoàn' };
  const vnParts = ms => { const d = new Date(ms + 7 * 3600000); return { d: pad(d.getUTCDate()), m: pad(d.getUTCMonth() + 1), y: d.getUTCFullYear(), h: pad(d.getUTCHours()), n: pad(d.getUTCMinutes()), s: pad(d.getUTCSeconds()) }; };
  const msDMY = ms => { if (!Number.isFinite(+ms) || +ms < 1e11) return ''; const t = vnParts(+ms); return `${t.d}/${t.m}/${t.y}`; };
  const msDMYHMS = ms => { if (!Number.isFinite(+ms) || +ms < 1e11) return ''; const t = vnParts(+ms); return `${t.d}/${t.m}/${t.y} ${t.h}:${t.n}:${t.s}`; };
  function apiToTables(items, empList) {
    // Mỗi dòng của API là một đơn như trên trang: KHÔNG loại dòng trùng mã hay mã rỗng (loại sẽ làm tổng đơn lệch so với trang).
    // Chỉ đếm số dòng trùng mã / mã rỗng để báo trong chẩn đoán khi tổng đơn không khớp.
    const seen = new Set(), rows = [], unknown = new Set(), codeById = new Map();
    let okN = 0, signN = 0, dupN = 0, emptyN = 0;
    for (const it of (items || [])) {
      if (it && it.dispDelivererId && it.delivererCode && !codeById.has(it.dispDelivererId)) codeById.set(it.dispDelivererId, it.delivererCode);
      const wb = String(it && it.billCode || '').trim().toUpperCase();
      if (!wb) emptyN++; else if (seen.has(wb)) dupN++; else seen.add(wb);
      const code = String(it.parcelStatus ?? '');
      let st = STATUS_TXT[code];
      if (!st) { st = `Trạng thái ${code}`; unknown.add(code); }
      const signMs = code === '4' ? signMsOf(it) : null;
      if (code === '4') { okN++; if (signMs !== null) signN++; }
      rows.push([wb, msDMY(it.predictTimeByDeparture), st, String(it.latestScanMan || '').trim(), msDMY(it.forwardAgingEndDate), signMs !== null ? msDMYHMS(signMs) : '']);
    }
    const orders = { headers: ['Mã vận đơn', 'Ngày cần phát hàng', 'Trạng thái đơn hàng', 'Người quét mới nhất', 'Thời hạn giao hàng nội bộ', 'Giờ giao thành công'], rows, wbCol: 0 };
    const eRows = (empList || []).map(e => {
      const c = codeById.get(e.dispDelivererId);
      return [`${String(e.delivererName || '').trim()}${c ? `(${c})` : ''}`, e.shouldDispCount ?? 0, e.todayDispCount ?? 0, msDMYHMS(e.firstDispScanTime), e.signCount ?? 0, String(e.departmentName ?? '').trim()];
    });
    const emp = { headers: ['Nhân viên giao nhận', 'Đơn cần phát', 'Đã quét phát', 'Quét phát đầu tiên', 'Đã ký nhận', 'Chức vụ'], rows: eRows };
    return { orders, emp, unknown: [...unknown], signKeys: timeKeysOf(items), okN, signN, dupN, emptyN };
  }

  // ---- Tính dashboard ----
  function computeDashboard({ orders, emp, tiktok, statusMap, now, rules }) {
    const H = orders.headers.map(h => String(h ?? '').normalize('NFC')), wb = orders.wbCol, rows = orders.rows;
    const stCol = colBy(H, /trạng thái/i);
    const scCol = colBy(H, /người quét|nhân viên|shipper|người phát|giao nhận/i);
    const dCol = colBy(H, /ngày cần phát/i);
    // Cột “Thời hạn giao hàng nội bộ”: ưu tiên tìm theo tên; nếu tiêu đề không đọc được thì suy ra từ cột ngày nằm ngay trước “Quá thời hạn giao hàng nội bộ”
    const looksDate = c => { const v = rows.map(r => r[c]).filter(x => String(x ?? '').trim()).slice(0, 60); return v.length >= 3 && v.filter(x => DATE_RE.test(String(x))).length / v.length >= 0.6; };
    let oCol = H.findIndex(h => /(thời\s*(hạn|gian)|hạn).*nội\s*bộ|odr/i.test(h) && !/quá|còn lại|chênh/i.test(h));
    if (oCol < 0) {
      const q = H.findIndex(h => /quá\s*thời\s*(hạn|gian).*nội\s*bộ/i.test(h));
      if (q > 0 && looksDate(q - 1)) oCol = q - 1;
    }
    // Nếu không có cột ngày thì dùng cột “Quá thời hạn giao hàng nội bộ” = 0 (đến hạn đúng hôm nay)
    let xCol = -1;
    if (oCol < 0) xCol = H.findIndex(h => /quá\s*thời\s*(hạn|gian).*nội\s*bộ/i.test(h));
    const tCol = colBy(H, /giờ giao thành công/i);
    const bl = { total: 0, wait: 0, fail: 0, unscanned: 0, d5: 0 };
    let odr = 0, firstOk = null;
    const odr2 = { due: 0, ok: 0 };
    const mkAgg = () => ({ total: 0, ok: 0, unscanned: 0, newTotal: 0, newOk: 0 });
    const all = mkAgg(), tt = mkAgg(), sme = mkAgg();
    const byCode = new Map(), byName = new Map(), stats = [];

    const getStat = cell => {
      const k = shipKeys(cell);
      if (!k.code && !k.name) return null;
      let o = k.code ? byCode.get(k.code) : byName.get(k.name);
      if (!o) { o = { label: String(cell).trim(), total: 0, ok: 0, ttTotal: 0, ttOk: 0 }; stats.push(o); }
      if (k.code) byCode.set(k.code, o);
      if (k.name && !byName.has(k.name)) byName.set(k.name, o);
      return o;
    };

    for (const r of rows) {
      const isTT = isTikTok(r[wb], tiktok);
      const cat = stCol >= 0 ? classify(r[stCol], statusMap) : 'other';
      const dateStr = dCol >= 0 ? r[dCol] : '';
      const isNew = dCol >= 0 && isToday(dateStr, now);
      const back = dCol >= 0 ? daysBefore(dateStr, now) : null;
      for (const a of (isTT ? [all, tt] : [all, sme])) {
        a.total++;
        if (cat === 'ok') a.ok++; else if (cat === 'unscanned') a.unscanned++;
        if (isNew) { a.newTotal++; if (cat === 'ok') a.newOk++; }
      }
      // Giờ đơn giao thành công đầu tiên trong ngày (đơn thành công có giờ ký nhận là hôm nay)
      if (tCol >= 0 && cat === 'ok' && isToday(r[tCol], now)) { const mm = minOfDay(r[tCol]); if (mm !== null && (firstOk === null || mm < firstOk)) firstOk = mm; }
      // Backlog: đơn có ngày cần phát cách hôm nay từ 3 ngày trở lên và chưa giao thành công (d5: từ 5 ngày)
      if (back !== null && back >= 3 && cat !== 'ok') {
        bl.total++;
        if (back >= 5) bl.d5++;
        const st = String(stCol >= 0 ? r[stCol] : '').normalize('NFC').toLowerCase();
        if (/chờ duyệt/.test(st)) bl.wait++;
        else if (/thất bại|không thành công/.test(st)) bl.fail++;
        else if (cat === 'unscanned') bl.unscanned++;
      }
      // TikTok đến hạn ODR: thời hạn giao hàng nội bộ là hôm nay
      if (isTT && oCol >= 0 && isToday(r[oCol], now)) { odr++; odr2.due++; if (cat === 'ok') odr2.ok++; }
      else if (isTT && oCol < 0 && xCol >= 0) { const m = String(r[xCol] ?? '').match(/\d+/); if (m && +m[0] === 0 && !/^-/.test(String(r[xCol]).trim())) { odr++; odr2.due++; if (cat === 'ok') odr2.ok++; } }
      // Đơn đã quét giao được tính cho người quét mới nhất
      if (cat !== 'unscanned' && scCol >= 0) {
        const o = getStat(r[scCol]);
        if (o) {
          o.total++;
          if (cat === 'ok') o.ok++;
          if (isTT) { o.ttTotal++; if (cat === 'ok') o.ttOk++; }
        }
      }
    }

    const W = workersOf(emp, now);
    const signIdx = W ? emp.headers.findIndex(h => h.includes('Đã ký nhận')) : -1;
    const ftIdx = W ? emp.headers.findIndex(h => h.includes('Quét phát đầu tiên')) : -1;
    const workersEstimated = !W;
    let nWork = W ? W.list.length : stats.filter(o => o.total > 0).length;   // với W: tính lại sau khi lọc ship đủ điều kiện
    const DM = dmRulesOf(rules), dmWhy = { dept: 0, late: 0, few: 0 };

    const used = new Set();
    let shippers;
    if (W) {
      shippers = W.list.map(row => {
        const cell = row[W.nameCol];
        const k = shipKeys(cell);
        const o = (k.code && byCode.get(k.code)) || (k.name && byName.get(k.name)) || null;
        if (o) used.add(o);
        const done = W.doneCol >= 0 ? parseNum(row[W.doneCol]) : NaN;
        return { name: cell, matched: !!o, scanned: Number.isFinite(done) ? done : (o ? o.total : 0),
          // GTC của shipper: lấy theo cột “Đã ký nhận” của trang (đúng số trang hiển thị); không có cột đó mới đếm từ bảng đơn
          ok: (signIdx >= 0 && Number.isFinite(parseNum(row[signIdx]))) ? parseNum(row[signIdx]) : (o ? o.ok : 0),
          ttTotal: o ? o.ttTotal : 0, ttOk: o ? o.ttOk : 0,
          signed: signIdx >= 0 ? parseNum(row[signIdx]) : NaN,
          ft: (ftIdx >= 0 && isToday(row[ftIdx], now)) ? minOfDay(row[ftIdx]) : null, _row: row };
      });
      // Chỉ giữ ship đủ điều kiện đi làm; ship không đủ bị ẩn (đơn của họ vẫn nằm trong các thẻ tổng)
      if (DM.on) {
        const hasDept = W.deptCol >= 0 && W.list.some(r => String(r[W.deptCol] ?? '').trim() !== '');   // bảng không có chức vụ thì không lọc theo chức vụ
        const allow = new Set(DM.depts.map(normDept)), cut = hmToMin(DM.cutoff), fin = dmFinal(now);
        shippers = shippers.filter(s => {
          if (hasDept && allow.size && !allow.has(normDept(s._row[W.deptCol]))) { dmWhy.dept++; return false; }
          if (Number.isFinite(cut) && Number.isFinite(s.ft) && s.ft > cut) { dmWhy.late++; return false; }
          if (fin && DM.minOk > 0 && !((Number.isFinite(s.ok) ? s.ok : 0) >= DM.minOk)) { dmWhy.few++; return false; }
          return true;
        });
      }
      shippers.forEach(s => { delete s._row; });
      nWork = shippers.length;
    } else {
      shippers = stats.filter(o => o.total > 0).map(o => ({ name: o.label, matched: true, scanned: o.total,
        ok: o.ok, ttTotal: o.ttTotal, ttOk: o.ttOk }));
    }
    shippers.forEach(s => { s.rateTrip = ratio(s.ok, s.scanned); s.rateTT = ratio(s.ttOk, s.ttTotal); });

    const sum = k => shippers.reduce((a, s) => a + (Number.isFinite(s[k]) ? s[k] : 0), 0);
    const totals = { name: 'TỔNG', scanned: sum('scanned'), ok: sum('ok'), ttTotal: sum('ttTotal'), ttOk: sum('ttOk') };
    totals.rateTrip = ratio(totals.ok, totals.scanned);
    totals.rateTT = ratio(totals.ttOk, totals.ttTotal);

    const notes = [];
    if (stCol < 0) notes.push('Không thấy cột “Trạng thái” trong bảng đơn nên chưa phân loại được thành công/chưa quét.');
    if (dCol < 0) notes.push('Không thấy cột “Ngày cần phát hàng” trong bảng đơn nên chưa tính được POD2400 và Backlog 3 ngày.');
    if (oCol < 0 && xCol < 0) notes.push('Không thấy cột “Thời hạn giao hàng nội bộ” trong bảng đơn nên chưa tính được TikTok đến hạn ODR.');
    if (scCol < 0) notes.push('Không thấy cột “Người quét” trong bảng đơn nên chưa tách được theo shipper.');
    if (W && W.note) notes.push(W.note);
    if (W && DM.on) {
      const hid = W.list.length - shippers.length;
      if (!(W.deptCol >= 0 && W.list.some(r => String(r[W.deptCol] ?? '').trim() !== ''))) notes.push('Không thấy cột “Chức vụ” trong bảng nhân viên nên chưa lọc theo chức vụ.');
      if (hid > 0) notes.push(`Đã ẩn ${hid}/${W.list.length} ship không đủ điều kiện đi làm (${[dmWhy.dept ? `${dmWhy.dept} sai chức vụ` : '', dmWhy.late ? `${dmWhy.late} quét đầu sau ${DM.cutoff}` : '', dmWhy.few ? `${dmWhy.few} giao thành công dưới ${DM.minOk} đơn (cuối ngày)` : ''].filter(Boolean).join(', ')}).`);
    }
    if (!W) notes.push('Chưa có bảng Tiến độ nhân viên: số nhân viên và danh sách shipper được ước tính từ bảng đơn.');
    const unmatched = shippers.filter(s => !s.matched).length;
    if (W && unmatched) notes.push(`${unmatched} shipper đi làm không khớp được tên/mã với bảng đơn (GTC tính bằng 0).`);
    const loose = stats.filter(o => !used.has(o)).reduce((a, o) => a + o.total, 0);
    if (W && loose) notes.push(`${fmt(loose)} đơn do người không có trong danh sách đi làm quét nên không nằm trong bảng shipper (vẫn tính trong các thẻ).`);
    if (W) {
      const signCol = emp.headers.findIndex(h => h.includes('Đã ký nhận'));
      if (signCol >= 0) {
        const empSigned = W.list.reduce((a, r) => { const n = parseNum(r[signCol]); return a + (Number.isFinite(n) ? n : 0); }, 0);
        const fromOrders = all.ok, gap = Math.abs(empSigned - fromOrders);
        notes.push(gap === 0
          ? `✔ GTC từng shipper lấy theo cột “Đã ký nhận”: tổng ${fmt(empSigned)}, khớp số đơn thành công đếm từ bảng đơn.`
          : gap <= Math.max(3, Math.round(fromOrders * 0.02))
            ? `GTC từng shipper lấy theo cột “Đã ký nhận”: tổng ${fmt(empSigned)} (đếm từ bảng đơn được ${fmt(fromOrders)}, lệch ${fmt(gap)} đơn là bình thường).`
            : `⚠ Cột “Đã ký nhận” ghi tổng ${fmt(empSigned)}, đếm từ bảng đơn được ${fmt(fromOrders)} (lệch ${fmt(gap)}). Kiểm tra phân loại trạng thái ở Cài đặt.`);
      }
    }

    const card = a => ({
      pod: ratio(a.newOk, a.newTotal), newOk: a.newOk, newTotal: a.newTotal,
      total: a.total, unscanned: a.unscanned, scanned: a.total - a.unscanned, ok: a.ok,
      scanRate: ratio(a.total - a.unscanned, a.total),
      perNV: nWork > 0 ? a.total / nWork : null,
      gtcTrip: ratio(a.ok, a.total - a.unscanned),
      gtcTotal: ratio(a.ok, a.total)
    });
    return {
      cards: { all: card(all), tt: card(tt), sme: card(sme) },
      extra: { backlog: dCol >= 0 ? bl : null, odrTT: (oCol >= 0 || xCol >= 0) ? odr : null, odr2: (oCol >= 0 || xCol >= 0) ? odr2 : null },
      firstOk, hasTime: tCol >= 0,
      hasTT: tt.total > 0, nWork, workersEstimated, W, shippers, totals, notes, stCol, scCol, dmHidden: W ? W.list.length - shippers.length : 0
    };
  }

  // Thẻ chỉ số: [nhãn, giá trị, kiểu, màu]; thẻ không có dữ liệu bị bỏ khi hiển thị
  function cardList(c) {
    const a = c.cards.all, t = c.cards.tt, s = c.cards.sme, e = c.extra || {};
    const bl = e.backlog;
    return {
      main: [
        ['Tổng đơn', a.total, 'int', 'blue'], ['Tỷ lệ quét giao', a.scanRate, 'pct', 'green'],
        ['POD2400', a.pod, 'pct', 'amber', [['Giao thành công', a.newOk], ['Tổng đơn mới hôm nay', a.newTotal]]],
        ['Chưa quét giao', a.unscanned, 'int', 'purple', [['SME', s.unscanned], ['TikTok', t.unscanned]]],
        ['GTC chuyến', a.gtcTrip, 'pct', 'teal'], ['GTC Total', a.gtcTotal, 'pct', 'orange'],
        ['Backlog 3 ngày', bl ? bl.total : null, 'int', 'red',
          bl ? [['Chờ duyệt', bl.wait], ['Giao thất bại', bl.fail], ['Nhập kho chưa quét phát', bl.unscanned]] : null]
      ],
      tts: [
        ['Tổng đơn TTS', t.total, 'int', 'pink'], ['Tỷ lệ quét TTS', t.scanRate, 'pct', 'green'],
        ['TikTok POD2400', t.pod, 'pct', 'amber', [['Giao thành công', t.newOk], ['Tổng đơn mới hôm nay', t.newTotal]]],
        ['TTS chưa quét giao', t.unscanned, 'int', 'purple'],
        ['GTC chuyến TTS', t.gtcTrip, 'pct', 'teal'], ['GTC Total TTS', t.gtcTotal, 'pct', 'orange'],
        ['TikTok đến hạn ODR', e.odrTT, 'int', 'red']
      ]
    };
  }

  function shipperCols(c) {
    const cols = [
      { k: 'name', label: 'Mã NV - Tên NV', group: '', kind: 'text' },
      { k: 'scanned', label: 'Đã quét giao', group: 'gh', kind: 'int', tip: 'Số đơn shipper đã quét giao' },
      { k: 'ok', label: 'GTC', group: 'gh', kind: 'int', tip: 'Số đơn giao thành công' },
      { k: 'rateTrip', label: '% GTC chuyến', group: 'gh', kind: 'pct', good: 'high', tip: 'GTC / Đã quét giao' }
    ];
    if (c.hasTT) {
      cols.push({ k: 'ttTotal', label: 'Tổng đơn TTS', group: 'tt', kind: 'int', tip: 'Số đơn TikTok shipper đã quét giao' });
      cols.push({ k: 'ttOk', label: 'GTC TTS', group: 'tt', kind: 'int', tip: 'Đơn TikTok giao thành công' });
      cols.push({ k: 'rateTT', label: '% GTC TTS', group: 'tt', kind: 'pct', good: 'high', tip: 'GTC TTS / Tổng đơn TTS' });
    }
    return cols;
  }

  function sortRows(rows, key, dir) {
    if (!key) return rows;
    const isN = v => typeof v === 'number' && Number.isFinite(v);
    return [...rows].sort((a, b) => {
      const x = a[key], y = b[key];
      if (isN(x) && isN(y)) return (x - y) * dir;
      if (isN(x)) return -1;
      if (isN(y)) return 1;
      return String(x ?? '').localeCompare(String(y ?? ''), 'vi') * dir;
    });
  }

  // ---- Lịch sử theo ngày ----
  const OFFICES = [];   // bản chia sẻ: không có bưu cục sẵn, thêm ở Cài đặt
  const WEEKDAY = ['CN', 'T2', 'T3', 'T4', 'T5', 'T6', 'T7'];
  const sameOffice = (a, b) => { const x = plain(a), y = plain(b); return !!x && !!y && (x === y || x.includes(y) || y.includes(x)); };
  const canonOffice = (name, list) => { const n = String(name ?? '').trim(); return (list || []).find(o => sameOffice(o, n)) || n; };

  const dateKey = d => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  const keyToUTC = k => { const [y, m, d] = String(k).split('-').map(Number); return Date.UTC(y, m - 1, d); };
  const utcToKey = t => { const d = new Date(t); return `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`; };
  const addDays = (k, n) => utcToKey(keyToUTC(k) + n * 86400000);
  const fmtDM = k => `${k.slice(8, 10)}/${k.slice(5, 7)}`;
  const fmtDMY = k => `${fmtDM(k)}/${k.slice(0, 4)}`;
  const fmtAvg = v => Number.isInteger(v) ? fmt(v) : round1(v).toLocaleString('vi-VN');

  // Tuần ISO (thứ Hai là ngày đầu tuần)
  function weekOf(k) {
    const t = keyToUTC(k), dow = (new Date(t).getUTCDay() + 6) % 7;
    const start = utcToKey(t - dow * 86400000), end = addDays(start, 6);
    const thu = keyToUTC(start) + 3 * 86400000, y = new Date(thu).getUTCFullYear();
    const doy = Math.round((thu - Date.UTC(y, 0, 1)) / 86400000) + 1, wk = Math.ceil(doy / 7);
    return { key: `${y}-W${pad(wk)}`, start, end, label: `Tuần ${wk} (${fmtDM(start)}–${fmtDM(end)}/${end.slice(0, 4)})` };
  }
  function monthBounds(k) {
    const y = +k.slice(0, 4), m = +k.slice(5, 7);
    return { first: `${k.slice(0, 7)}-01`, last: addDays(utcToKey(Date.UTC(y, m, 1)), -1) };
  }
  function periodOf(k, mode) {
    if (mode === 'week') return weekOf(k);
    if (mode === 'month') return { key: k.slice(0, 7), label: `Tháng ${Number(k.slice(5, 7))}/${k.slice(0, 4)}` };
    return { key: k, label: `${WEEKDAY[new Date(keyToUTC(k)).getUTCDay()]} ${fmtDMY(k)}` };
  }
  function rangeFor(preset, today) {
    switch (preset) {
      case 'd7': return { from: addDays(today, -6), to: today };
      case 'd30': return { from: addDays(today, -29), to: today };
      case 'wk': { const w = weekOf(today); return { from: w.start, to: w.end }; }
      case 'lastwk': { const w = weekOf(addDays(today, -7)); return { from: w.start, to: w.end }; }
      case 'mo': { const b = monthBounds(today); return { from: b.first, to: b.last }; }
      case 'lastmo': { const b = monthBounds(addDays(monthBounds(today).first, -1)); return { from: b.first, to: b.last }; }
      default: return { from: '0000-01-01', to: '9999-12-31' };
    }
  }

  const nz = v => (typeof v === 'number' && Number.isFinite(v)) ? v : 0;
  const nn = v => (typeof v === 'number' && Number.isFinite(v)) ? v : null;
  const cntOf = c => { c = c || {}; return { total: nz(c.total), unscanned: nz(c.unscanned), ok: nz(c.ok), newTotal: nz(c.newTotal), newOk: nz(c.newOk) }; };
  const blOf = b => b ? { total: nz(b.total), wait: nz(b.wait), fail: nz(b.fail), unscanned: nz(b.unscanned), d5: nn(b.d5) } : null;
  const shipsOf = list => (Array.isArray(list) ? list : []).map(s => ({
    n: String(s.n ?? s.name ?? '').trim().slice(0, 80), sc: nz(s.sc ?? s.scanned), ok: nz(s.ok),
    tt: nz(s.tt ?? s.ttTotal), tk: nz(s.tk ?? s.ttOk), sg: nn(s.sg ?? s.signed), ft: nn(s.ft)
  })).filter(s => s.n);

  // Làm sạch một bản ghi (dùng khi đọc từ kho, khi nạp file sao lưu); trả về null nếu không hợp lệ
  function normalizeRecord(r) {
    if (!r || typeof r !== 'object') return null;
    const office = String(r.office ?? '').trim(), date = String(r.date ?? '');
    if (!office || !/^\d{4}-\d{2}-\d{2}$/.test(date)) return null;
    const ships = shipsOf(r.ships);
    // Số shipper có quét giao hàng (đã quét phát ít nhất 1 đơn); bản ghi cũ chưa có thì suy ra từ danh sách shipper
    const nScan = nn(r.nScan) ?? (ships.length ? ships.filter(x => x.sc > 0).length : null);
    return {
      id: `${office}|${date}`, office, date, savedAt: String(r.savedAt ?? ''), v: 1, verified: !!r.verified,
      all: cntOf(r.all), tt: cntOf(r.tt), backlog: blOf(r.backlog), odrTT: nn(r.odrTT),
      odr2: (r.odr2 && typeof r.odr2 === 'object') ? { due: nz(r.odr2.due), ok: nz(r.odr2.ok) } : null, firstOk: nn(r.firstOk),
      nWork: nz(r.nWork), nScan, workersEstimated: !!r.workersEstimated, ships
    };
  }

  // Bản chụp chỉ số của một bưu cục trong một ngày (không có mã vận đơn, không có thông tin người nhận)
  function buildRecord({ office, now, calc, verified }) {
    const e = calc.extra || {};
    return normalizeRecord({
      office, date: dateKey(now), savedAt: now.toISOString(), verified,
      all: calc.cards.all, tt: calc.cards.tt, backlog: e.backlog || null, odrTT: e.odrTT, odr2: e.odr2 || null, firstOk: calc.firstOk ?? null,
      nWork: calc.nWork, workersEstimated: calc.workersEstimated, ships: calc.shippers
    });
  }

  // Chi tiết kèm thông báo lệch tổng đơn: số dòng API trả, số trang đã gọi, dòng trùng mã, dòng mã rỗng
  function diagText(d) {
    if (!d || typeof d !== 'object') return '';
    const p = [];
    if (Number.isFinite(d.rows)) p.push(`API trả ${d.rows} dòng${Number.isFinite(d.pages) ? ` qua ${d.pages} lần gọi` : ''}`);
    if (Number.isFinite(d.dup)) p.push(`${d.dup} dòng trùng mã`);
    if (Number.isFinite(d.empty)) p.push(`${d.empty} dòng mã rỗng`);
    if (Number.isFinite(d.dropped) && d.dropped > 0) p.push(`bỏ ${d.dropped} dòng giống hệt khi gộp trang`);
    return p.length ? ` (${p.join('; ')})` : '';
  }
  // Có nên tự lưu không: đúng bưu cục và tổng đơn khớp trang
  function decideSave(x) {
    if (!x.hasOrders) return { save: false, force: false, reason: 'Chưa đọc được bảng chi tiết đơn.' };
    const page = String(x.pageOffice ?? '').trim();
    if (!page) return { save: false, force: false, reason: 'Không đọc được tên bưu cục trên trang nên không biết lưu vào bưu cục nào.' };
    if (x.startOffice && !sameOffice(x.startOffice, page))
      return { save: false, force: true, reason: `Bưu cục trên trang đổi từ “${x.startOffice}” sang “${page}” trong lúc tải dữ liệu.` };
    if (x.expectedOffice && !sameOffice(x.expectedOffice, page))
      return { save: false, force: true, reason: `Trang đang ở bưu cục “${page}”, không phải “${x.expectedOffice}” bạn đã chọn.` };
    const hasOv = Number.isFinite(x.ovNeed);
    if (hasOv && x.ovNeed !== x.computedTotal)
      return { save: false, force: true, reason: `Tổng đơn tự đếm ${x.computedTotal} khác số trên trang ${x.ovNeed}${diagText(x.diag)}.` };
    return { save: true, verified: hasOv };
  }

  // Gộp nhiều bản ghi (nhiều ngày và/hoặc nhiều bưu cục): cộng số đếm, rồi tính lại tỷ lệ từ tổng
  function aggregate(recs) {
    const dates = [...new Set(recs.map(r => r.date))].sort(), days = dates.length;
    const A = { total: 0, unscanned: 0, ok: 0, newTotal: 0, newOk: 0 }, T = { ...A };
    let bl = 0, blN = 0, od = 0, odN = 0, wk = 0, sn = 0, snN = 0, b5 = 0, b5N = 0, o2d = 0, o2o = 0, o2N = 0, fo = null;
    for (const r of recs) {
      for (const k of Object.keys(A)) { A[k] += nz(r.all[k]); T[k] += nz(r.tt[k]); }
      if (r.backlog) { bl += r.backlog.total; blN++; if (Number.isFinite(r.backlog.d5)) { b5 += r.backlog.d5; b5N++; } }
      if (r.odr2) { o2d += r.odr2.due; o2o += r.odr2.ok; o2N++; }
      if (Number.isFinite(r.firstOk)) fo = fo === null ? r.firstOk : Math.min(fo, r.firstOk);
      if (Number.isFinite(r.odrTT)) { od += r.odrTT; odN++; }
      wk += nz(r.nWork);
      if (Number.isFinite(r.nScan)) { sn += r.nScan; snN++; }
    }
    const lastDate = dates[days - 1] || '', last = recs.filter(r => r.date === lastDate);
    const lastBl = last.some(r => r.backlog) ? last.reduce((a, r) => a + (r.backlog ? r.backlog.total : 0), 0) : null;
    const lastOd = last.some(r => Number.isFinite(r.odrTT)) ? last.reduce((a, r) => a + (Number.isFinite(r.odrTT) ? r.odrTT : 0), 0) : null;
    return { days, A, T, bl: blN ? bl : null, od: odN ? od : null, wk, sn: snN ? sn : null, lastBl, lastOd, firstDate: dates[0] || '', lastDate, b5: b5N ? b5 : null, od2: o2N ? [o2o, o2d] : null, fo };
  }
  function metricsOf(g) {
    const d = g.days || 0, A = g.A, T = g.T;
    const avg = v => (v === null || v === undefined || !d) ? null : v / d;
    return {
      days: d, total: avg(A.total), unscanned: avg(A.unscanned),
      scanRate: ratio(A.total - A.unscanned, A.total),
      pod: ratio(A.newOk, A.newTotal), podN: [A.newOk, A.newTotal],
      podTT: ratio(T.newOk, T.newTotal), podTTN: [T.newOk, T.newTotal],
      gtcTrip: ratio(A.ok, A.total - A.unscanned), gtcTotal: ratio(A.ok, A.total), gtcN: [A.ok, A.total],
      backlog: avg(g.bl), backlogLast: g.lastBl, odr: avg(g.od), odrLast: g.lastOd, nWork: avg(g.wk), nScan: avg(g.sn),
      backlog5: avg(g.b5), odrPct: g.od2 ? ratio(g.od2[0], g.od2[1]) : null, odrPctN: g.od2, firstOk: g.fo
    };
  }
  function groupRecords(recs, mode) {
    const map = new Map();
    for (const r of [...recs].sort((a, b) => a.date < b.date ? -1 : a.date > b.date ? 1 : 0)) {
      const p = periodOf(r.date, mode);
      let g = map.get(p.key);
      if (!g) { g = { key: p.key, label: p.label, recs: [] }; map.set(p.key, g); }
      g.recs.push(r);
    }
    return [...map.values()].map(g => { g.agg = aggregate(g.recs); g.m = metricsOf(g.agg); return g; });
  }
  // Mức tăng/giảm (điểm %) so với kỳ có dữ liệu liền trước; list sắp xếp từ cũ đến mới
  function withDeltas(list) {
    const keys = ['scanRate', 'pod', 'podTT', 'gtcTrip', 'gtcTotal'];
    list.forEach((g, i) => {
      g.delta = {};
      if (!i) return;
      const p = list[i - 1].m;
      for (const k of keys) if (Number.isFinite(g.m[k]) && Number.isFinite(p[k])) g.delta[k] = g.m[k] - p[k];
    });
    return list;
  }

  // Gộp shipper theo mã (hoặc tên) qua nhiều ngày; recs phải sắp xếp từ cũ đến mới
  function aggregateShippers(recs) {
    const map = new Map();
    for (const r of recs) for (const s of r.ships) {
      const k = shipKeys(s.n), key = k.code || k.name || s.n;
      let o = map.get(key);
      if (!o) { o = { name: s.n, offices: new Set(), dates: new Set(), sc: 0, ok: 0, tt: 0, tk: 0, sg: 0, sgN: 0 }; map.set(key, o); }
      o.name = s.n; o.offices.add(r.office); o.dates.add(r.date);
      o.sc += s.sc; o.ok += s.ok; o.tt += s.tt; o.tk += s.tk;
      if (s.sg !== null) { o.sg += s.sg; o.sgN++; }
    }
    return [...map.entries()].map(([key, o]) => ({
      key, name: o.name, office: [...o.offices].join(', '), days: o.dates.size,
      sc: o.sc, avgSc: o.dates.size ? o.sc / o.dates.size : null, ok: o.ok, rateTrip: ratio(o.ok, o.sc),
      tt: o.tt, tk: o.tk, rateTT: ratio(o.tk, o.tt), sg: o.sgN ? o.sg : null
    }));
  }

  // Sao lưu / khôi phục
  function validateBackup(obj) {
    if (!obj || typeof obj !== 'object' || obj.app !== '800best-history' || !Array.isArray(obj.records))
      return { ok: false, records: [], error: 'File không phải bản sao lưu lịch sử 800BEST.' };
    const records = obj.records.map(normalizeRecord).filter(Boolean);
    if (!records.length) return { ok: false, records: [], error: 'File sao lưu không có bản ghi hợp lệ.' };
    return { ok: true, records, skipped: obj.records.length - records.length };
  }
  function planMerge(existing, incoming) {
    const put = []; let added = 0, updated = 0, kept = 0;
    for (const r of incoming) {
      const cur = existing.get(r.id);
      if (!cur) { put.push(r); added++; }
      else if (String(r.savedAt) > String(cur.savedAt)) { put.push(r); updated++; }
      else kept++;
    }
    return { put, added, updated, kept };
  }
  const DAY_MS = 86400000;
  function backupDue(count, lastMs, nowMs) { return count >= 3 && (!lastMs || nowMs - lastMs > 7 * DAY_MS); }

  /* ---- ANM: so sánh nhiều bưu cục ---- */
  const lastDayOfMonth = (y, m) => new Date(Date.UTC(y, m, 0)).getUTCDate(); // m = 1..12
  function sameDayPrevMonth(k) {
    let y = +k.slice(0, 4), m = +k.slice(5, 7) - 1; const d = +k.slice(8, 10);
    if (!m) { m = 12; y--; }
    return `${y}-${pad(m)}-${pad(Math.min(d, lastDayOfMonth(y, m)))}`;
  }
  // Ngày gần target nhất (trong ±3 ngày) có cùng thứ với ref
  function nearestSameWeekday(target, ref) {
    const dt = (new Date(keyToUTC(ref)).getUTCDay() - new Date(keyToUTC(target)).getUTCDay() + 7) % 7;
    return addDays(target, dt > 3 ? dt - 7 : dt);
  }
  const CMP_MODES = ['slot2', 'yd', 'wk1', 'mo_dom', 'mo_dow', 'week', 'month'];
  function cmpPeriods(base, mode) {
    const day = k => ({ from: k, to: k, label: periodOf(k, 'day').label });
    switch (mode) {
      case 'slot2': case 'yd': return { cur: day(base), prev: day(addDays(base, -1)) };
      case 'mo_dom': return { cur: day(base), prev: day(sameDayPrevMonth(base)) };
      case 'mo_dow': return { cur: day(base), prev: day(nearestSameWeekday(sameDayPrevMonth(base), base)) };
      case 'week': {
        const w = weekOf(base), ps = addDays(w.start, -7), pe = addDays(base, -7);
        return { cur: { from: w.start, to: base, label: `Tuần này ${fmtDM(w.start)}–${fmtDM(base)}` }, prev: { from: ps, to: pe, label: `Tuần trước ${fmtDM(ps)}–${fmtDM(pe)}` } };
      }
      case 'month': {
        const pe = sameDayPrevMonth(base), first = `${base.slice(0, 7)}-01`, pf = `${pe.slice(0, 7)}-01`;
        return {
          cur: { from: first, to: base, label: `Tháng ${+base.slice(5, 7)}/${base.slice(0, 4)} (1–${+base.slice(8, 10)})` },
          prev: { from: pf, to: pe, label: `Tháng ${+pe.slice(5, 7)}/${pe.slice(0, 4)} (1–${+pe.slice(8, 10)})` }
        };
      }
      default: return { cur: day(base), prev: day(addDays(base, -7)) };
    }
  }

  const ANM_METRICS = [
    { k: 'total', label: 'Tổng đơn', group: 'Sản lượng', kind: 'avg' },
    { k: 'perNV', label: 'Đơn/NV', group: 'Sản lượng', kind: 'avg' },
    { k: 'nWork', label: 'DM đi làm', group: 'Sản lượng', kind: 'avg' },
    { k: 'nScan', label: 'DM có quét giao', group: 'Sản lượng', kind: 'avg', high: true },
    { k: 'scanRate', label: 'Tỷ lệ quét giao', group: 'Quét giao', kind: 'pct' },
    { k: 'unscanned', label: 'Đơn chưa quét', group: 'Quét giao', kind: 'avg', low: true },
    { k: 'pod', label: 'POD2400', group: 'Giao thành công', kind: 'pct' },
    { k: 'gtcTotal', label: 'GTC Total', group: 'Giao thành công', kind: 'pct' },
    { k: 'backlog', label: 'Backlog 3 ngày', group: 'Tồn đọng', kind: 'avg', low: true },
    { k: 'podTT', label: 'POD TTS', group: 'TikTok Shop', kind: 'pct' },
    { k: 'gtcTT', label: 'GTC TTS', group: 'TikTok Shop', kind: 'pct' },
    { k: 'odr', label: 'TTS đến hạn ODR', group: 'TikTok Shop', kind: 'avg', low: true },
    { k: 'odrPct', label: 'Tỷ lệ ODR', group: 'TikTok Shop', kind: 'pct' },
    { k: 'backlog5', label: 'Backlog 5 ngày', group: 'Tồn đọng', kind: 'avg', low: true }
  ];
  // Bốn tiêu chí xếp hạng thi đua giữa các bưu cục
  const RANK_KEYS = ['pod', 'podTT', 'gtcTotal', 'odrPct'];
  const RANK_LABEL = { pod: 'POD2400', podTT: 'POD2400 TikTok', gtcTotal: 'Giao thành công', odrPct: 'Tỷ lệ ODR' };
  // Ngưỡng mặc định: chỉ số tỷ lệ (cao là tốt): dưới lo = đỏ, từ hi = xanh; chỉ số đếm (thấp là tốt): ≤ ok xanh, > warn đỏ
  const ANM_TH = {
    scanRate: { lo: 85, hi: 92 }, pod: { lo: 60, hi: 75 }, gtcTotal: { lo: 70, hi: 80 },
    podTT: { lo: 60, hi: 70 }, gtcTT: { lo: 60, hi: 70 },
    backlog: { ok: 50, warn: 100 }, odr: { ok: 20, warn: 40 }, odrPct: { lo: 70, hi: 85 }, backlog5: { ok: 30, warn: 80 }
  };
  function levelOf(k, v, user) {
    const t = Object.assign({}, ANM_TH[k], user && user[k]);
    if (!ANM_TH[k] || !Number.isFinite(v)) return '';
    if (t.lo !== undefined) return v >= t.hi ? 'g' : v >= t.lo ? 'a' : 'r';
    return v <= t.ok ? 'g' : v <= t.warn ? 'a' : 'r';
  }
  // Chỉ số của một nhóm bản ghi (một hoặc nhiều ngày/bưu cục); null nếu không có bản ghi
  function anmMetrics(recs) {
    if (!recs || !recs.length) return null;
    const g = aggregate(recs), m = metricsOf(g);
    m.perNV = g.wk > 0 ? g.A.total / g.wk : null;
    m.gtcTT = ratio(g.T.ok, g.T.total);
    m.n = recs.length;
    return m;
  }
  function buildAlerts(rows, user) {
    const out = [], lab = k => ANM_METRICS.find(x => x.k === k).label;
    for (const r of rows) {
      if (!r.cur) { out.push({ level: 'a', office: r.office, text: r.noData || 'Chưa có số liệu trong kỳ đang xem (cần tải bưu cục này).' }); continue; }
      const red = new Set();
      for (const k of ['pod', 'gtcTotal', 'scanRate', 'podTT', 'gtcTT']) {
        const v = r.cur[k];
        if (Number.isFinite(v) && levelOf(k, v, user) === 'r') {
          red.add(k);
          out.push({ level: 'r', office: r.office, text: `${lab(k)} chỉ ${fmtVal(v, 'pct')} (dưới ${Object.assign({}, ANM_TH[k], user && user[k]).lo}%)` });
        }
      }
      const bl = r.cur.backlog, od = r.cur.odr;
      if (Number.isFinite(bl) && levelOf('backlog', bl, user) === 'r') { red.add('backlog'); out.push({ level: 'r', office: r.office, text: `Backlog 3 ngày ${fmtAvg(bl)} đơn (trên ${Object.assign({}, ANM_TH.backlog, user && user.backlog).warn})` }); }
      if (Number.isFinite(od) && levelOf('odr', od, user) === 'r') out.push({ level: 'r', office: r.office, text: `${fmtAvg(od)} đơn TikTok đến hạn ODR (trên ${Object.assign({}, ANM_TH.odr, user && user.odr).warn})` });
      if (r.prev) {
        for (const k of ['pod', 'gtcTotal', 'scanRate']) {
          const d = r.cur[k] - r.prev[k];
          if (!red.has(k) && Number.isFinite(d) && d <= -3) out.push({ level: 'a', office: r.office, text: `${lab(k)} giảm ${round1(-d).toLocaleString('vi-VN')} điểm so với kỳ trước` });
        }
        const db = r.cur.backlog - r.prev.backlog;
        if (!red.has('backlog') && Number.isFinite(db) && db >= 20) out.push({ level: 'a', office: r.office, text: `Backlog tăng ${fmtAvg(db)} đơn so với kỳ trước` });
      }
      if (r.recs && r.recs.some(x => !x.verified)) out.push({ level: 'a', office: r.office, text: 'Có số liệu chưa đối chiếu được với trang gốc (dấu ⚠ trong Lịch sử).' });
    }
    const merged = [];
    for (const lv of ['r', 'a']) for (const r of rows) {
      const t = out.filter(x => x.level === lv && x.office === r.office).map(x => x.text);
      if (t.length) merged.push({ level: lv, office: r.office, text: t.join(' · ') });
    }
    return merged;
  }

  /* ---- Bản chụp theo mốc giờ (tab ANM) ---- */
  const SNAP_SLOTS = ['10:30', '13:30', '15:30', '18:30', '23:30'];
  const SNAP_WINDOW_MIN = 15, SNAP_KEEP_DAYS = 90, SNAP_RETRY_MS = 170000, SNAP_MAX_TRY = 5;
  const slotMin = sl => { const m = /^(\d{2}):(\d{2})$/.exec(String(sl)); return m && +m[1] < 24 && +m[2] < 60 ? +m[1] * 60 + +m[2] : NaN; };
  const minutesOf = d => d.getHours() * 60 + d.getMinutes() + d.getSeconds() / 60;
  // Mốc đang trong khung chạy (từ đúng giờ mốc đến +15 phút); trả về mốc muộn nhất nếu có
  function dueSlot(now) {
    const t = minutesOf(now); let hit = null;
    for (const sl of SNAP_SLOTS) { const m = slotMin(sl); if (t >= m && t < m + SNAP_WINDOW_MIN) hit = sl; }
    return hit;
  }
  // Mốc gần nhất đã đến giờ trong ngày (dùng để so số mới nhất với bản lưu cùng mốc)
  function lastSlotBefore(now) { const t = minutesOf(now); let r = null; for (const sl of SNAP_SLOTS) if (t >= slotMin(sl)) r = sl; return r; }
  function slotPhase(dateStr, slot, now) {
    const today = dateKey(now);
    if (dateStr < today) return 'past';
    if (dateStr > today) return 'future';
    const t = minutesOf(now), m = slotMin(slot);
    return t < m ? 'future' : t < m + SNAP_WINDOW_MIN ? 'window' : 'past';
  }
  // Bản chụp: cùng cấu trúc bản ghi ngày + mốc giờ (không có mã vận đơn, không có thông tin người nhận)
  function normalizeSnap(r) {
    const base = normalizeRecord(r), slot = String((r && r.slot) ?? '');
    if (!base || !Number.isFinite(slotMin(slot))) return null;
    // Bản chụp giữ danh sách ship dạng gọn (tên, số quét, số giao thành công, giờ quét đầu) để so sánh từng ship cùng mốc giờ
    const ships = base.ships.slice(0, 150).map(x => ({ n: x.n, sc: x.sc, ok: x.ok, tt: 0, tk: 0, sg: null, ft: x.ft }));
    return { ...base, id: `${base.office}|${base.date}|${slot}`, slot, runAt: String((r && r.runAt) ?? base.savedAt), ships };
  }
  function buildSnap({ office, now, calc, verified, slot }) {
    return normalizeSnap({ ...buildRecord({ office, now, calc, verified }), slot, runAt: now.toISOString() });
  }
  function snapsToIndex(list) {   // Map(bưu cục → Map('ngày|mốc' → bản chụp))
    const idx = new Map();
    for (const r of list) { if (!idx.has(r.office)) idx.set(r.office, new Map()); idx.get(r.office).set(`${r.date}|${r.slot}`, r); }
    return idx;
  }
  const snapAt = (idx, office, date, slot) => { const m = idx.get(office); return (m && m.get(`${date}|${slot}`)) || null; };
  function snapsToDelete(list, today, keepDays = SNAP_KEEP_DAYS) { const lim = addDays(today, -keepDays); return list.filter(r => r.date < lim).map(r => r.id); }
  // Tình trạng một mốc của một ngày: ok (đủ bưu cục) / part (thiếu) / due (đang trong khung chạy) / miss (lỡ) / future
  function slotState(dateStr, slot, now, offices, idx) {
    const have = offices.filter(o => snapAt(idx, o, dateStr, slot)).length, n = offices.length;
    if (n && have === n) return { st: 'ok', have, n };
    if (have) return { st: 'part', have, n };
    const ph = slotPhase(dateStr, slot, now);
    return { st: ph === 'future' ? 'future' : ph === 'window' ? 'due' : 'miss', have, n };
  }

  /* ---- Thời tiết (Open-Meteo, lượng mưa theo giờ) ---- */
  const WX_DEFAULT = { on: false, lat: 21.0285, lon: 105.8542 };   // bản chia sẻ: tắt sẵn, nhập tọa độ của bạn ở Cài đặt rồi bật
  const wxUrl = (lat, lon) => `https://api.open-meteo.com/v1/forecast?latitude=${encodeURIComponent(lat)}&longitude=${encodeURIComponent(lon)}&hourly=precipitation&past_days=14&forecast_days=1&timezone=Asia%2FBangkok`;
  function parseWeather(j) {
    const h = j && j.hourly;
    if (!h || !Array.isArray(h.time) || !Array.isArray(h.precipitation) || h.time.length !== h.precipitation.length || !h.time.length) return null;
    return { times: h.time.map(String), mm: h.precipitation.map(v => (typeof v === 'number' && Number.isFinite(v)) ? v : null) };
  }
  // Tổng lượng mưa (mm) của một ngày từ 06:00 đến trước phút uptoMin; null nếu không có số liệu giờ nào trong khoảng đó
  function rainSum(w, dayKey, uptoMin) {
    if (!w || !Array.isArray(w.times)) return null;
    let sum = 0, n = 0;
    for (let i = 0; i < w.times.length; i++) {
      const t = w.times[i];
      if (t.slice(0, 10) !== dayKey) continue;
      const m = +t.slice(11, 13) * 60;
      if (m >= 360 && m < uptoMin && w.mm[i] !== null) { sum += w.mm[i]; n++; }
    }
    return n ? Math.round(sum * 10) / 10 : null;
  }

  /* ---- Phân tích nguyên nhân POD2400 tăng/giảm so với hôm qua cùng mốc giờ ---- */
  const RS_TH = { minDelta: 0.5, shipPct: 10, vol: 15, load: 10, backlog: 20, backlogPct: 10, late: 30, first: 30, rain: 2, rainGap: 1.5, shipRate: 20, shipMin: 15, top: 4 };
  const shipLabel = n => (String(n ?? '').replace(/\s*\([^)]*\)\s*$/, '').trim() || String(n ?? '').trim());
  function shipMapOf(ships) {
    const m = new Map();
    for (const x of (ships || [])) { const k = shipKeys(x.n), key = k.code || k.name || x.n; if (!m.has(key)) m.set(key, x); }
    return m;
  }
  const pct0 = v => `${Math.round(v)}%`;
  // cur, yd: bản ghi (ngày hoặc bản chụp) của cùng một bưu cục; rain: { cur, yd } (mm) hoặc null
  function buildReasons({ cur, yd, rain }) {
    if (!cur || !yd) return { dir: 'none', delta: null, reasons: [], why: !cur ? 'Chưa có số liệu của kỳ đang xem.' : 'Chưa có bản lưu hôm qua cùng mốc giờ để so sánh.' };
    const pc = ratio(cur.all.newOk, cur.all.newTotal), py = ratio(yd.all.newOk, yd.all.newTotal);
    if (!Number.isFinite(pc) || !Number.isFinite(py)) return { dir: 'none', delta: null, reasons: [], why: 'Chưa tính được POD2400 của một trong hai ngày.' };
    const delta = pc - py;
    if (Math.abs(delta) < RS_TH.minDelta) return { dir: 'flat', delta, pc, py, reasons: [] };
    const sign = delta > 0 ? 1 : -1, out = [];
    const add = (s, icon, text, w, k = 1) => { if (s === sign) out.push({ icon, text, w: Math.min(5, w) * k }); };   // k: độ tin cậy của loại lý do (ship, sản lượng trực tiếp nhất)
    // 1. Số ship làm và đơn trên mỗi ship
    const useScan = Number.isFinite(cur.nScan) && Number.isFinite(yd.nScan);
    const nc = useScan ? cur.nScan : cur.nWork, ny = useScan ? yd.nScan : yd.nWork;
    const loadC = nc > 0 ? cur.all.newTotal / nc : null, loadY = ny > 0 ? yd.all.newTotal / ny : null;
    const loadPct = (loadC !== null && loadY !== null && loadY > 0) ? (loadC - loadY) / loadY * 100 : null;
    if (Number.isFinite(nc) && Number.isFinite(ny)) {
      const dN = nc - ny, rel = ny > 0 ? Math.abs(dN) / ny * 100 : 0;
      if (Math.abs(dN) >= 1 && rel >= RS_TH.shipPct) {
        const ext = (loadPct !== null && Math.abs(loadPct) >= RS_TH.load) ? `, mỗi ship ${loadPct > 0 ? 'gánh nhiều hơn' : 'nhẹ hơn'} ${pct0(Math.abs(loadPct))}` : '';
        add(dN > 0 ? 1 : -1, '🚚', `${fmt(nc)} ship ${useScan ? 'có quét giao' : 'đi làm'} (${dN > 0 ? '+' : '−'}${Math.abs(dN)} so với hôm qua)${ext}`, rel / RS_TH.shipPct, 1.6);
      } else if (loadPct !== null && Math.abs(loadPct) >= RS_TH.load) {
        add(loadPct > 0 ? -1 : 1, '🚚', `Đơn trên mỗi ship ${fmt(Math.round(loadC))} (hôm qua ${fmt(Math.round(loadY))}), ${loadPct > 0 ? 'nhiều hơn' : 'ít hơn'} ${pct0(Math.abs(loadPct))}`, Math.abs(loadPct) / RS_TH.load, 1);
      }
    }
    // 2. Sản lượng đơn mới
    if (yd.all.newTotal > 0) {
      const vp = (cur.all.newTotal - yd.all.newTotal) / yd.all.newTotal * 100;
      if (Math.abs(vp) >= RS_TH.vol) add(vp > 0 ? -1 : 1, '📦', `Sản lượng đơn mới ${vp > 0 ? 'tăng' : 'giảm'} ${pct0(Math.abs(vp))} (${fmt(cur.all.newTotal)} so với ${fmt(yd.all.newTotal)} đơn)`, Math.abs(vp) / RS_TH.vol, 1.3);
    }
    // 3. Backlog 3 ngày
    if (cur.backlog && yd.backlog && yd.backlog.total > 0) {
      const d = cur.backlog.total - yd.backlog.total, rel = Math.abs(d) / yd.backlog.total * 100;
      if (Math.abs(d) >= RS_TH.backlog && rel >= RS_TH.backlogPct) add(d > 0 ? -1 : 1, '🗂', `Backlog 3 ngày ${d > 0 ? 'tăng' : 'giảm'} ${fmt(Math.abs(d))} đơn (${fmt(cur.backlog.total)} so với ${fmt(yd.backlog.total)})`, rel / RS_TH.backlogPct, 0.5);
    }
    // 4. Giờ có đơn giao thành công đầu tiên
    if (Number.isFinite(cur.firstOk) && Number.isFinite(yd.firstOk)) {
      const d = cur.firstOk - yd.firstOk;
      if (Math.abs(d) >= RS_TH.first) add(d > 0 ? -1 : 1, '⏱', `Đơn giao thành công đầu tiên lúc ${hmOf(cur.firstOk)} (hôm qua ${hmOf(yd.firstOk)})`, Math.abs(d) / RS_TH.first, 0.6);
    }
    // 5. Từng ship: % giao thành công cùng mốc giờ so với chính họ hôm qua, và giờ ra đường
    const mc = shipMapOf(cur.ships), my = shipMapOf(yd.ships);
    const per = [], lates = [];
    for (const [k, a] of mc) {
      const b = my.get(k); if (!b) continue;
      if (a.sc >= RS_TH.shipMin && b.sc >= RS_TH.shipMin) per.push({ name: shipLabel(a.n), rc: a.ok / a.sc * 100, ry: b.ok / b.sc * 100, sc: a.sc });
      if (Number.isFinite(a.ft) && Number.isFinite(b.ft)) lates.push(a.ft - b.ft);
    }
    per.forEach(x => { x.diff = x.rc - x.ry; });
    per.filter(x => (x.diff > 0 ? 1 : -1) === sign && Math.abs(x.diff) >= RS_TH.shipRate)
      .sort((a, b) => Math.abs(b.diff) * Math.sqrt(b.sc) - Math.abs(a.diff) * Math.sqrt(a.sc)).slice(0, 2)
      .forEach(x => add(sign, '👤', `Ship ${x.name} giao ${pct0(x.rc)} (hôm qua ${pct0(x.ry)}, cùng mốc giờ)`, Math.abs(x.diff) / RS_TH.shipRate, 0.9));
    if (lates.length >= 3) {
      const avg = lates.reduce((a, v) => a + v, 0) / lates.length;
      if (Math.abs(avg) >= RS_TH.late) add(avg > 0 ? -1 : 1, '🕘', `${lates.length} ship ra đường ${avg > 0 ? 'muộn hơn' : 'sớm hơn'} trung bình ${Math.round(Math.abs(avg))} phút so với hôm qua`, Math.abs(avg) / RS_TH.late, 0.7);
    }
    // 6. Thời tiết
    if (rain && Number.isFinite(rain.cur) && Number.isFinite(rain.yd)) {
      const f = v => v.toLocaleString('vi-VN');
      if (rain.cur >= RS_TH.rain && rain.cur - rain.yd >= RS_TH.rainGap) add(-1, '🌧', `Hôm nay mưa ${f(rain.cur)} mm (hôm qua ${f(rain.yd)} mm)`, (rain.cur - rain.yd) / RS_TH.rain, 0.9);
      else if (rain.yd >= RS_TH.rain && rain.yd - rain.cur >= RS_TH.rainGap) add(1, rain.cur < 0.5 ? '☀' : '🌦', `Hôm qua mưa ${f(rain.yd)} mm, hôm nay ${rain.cur < 0.5 ? 'không mưa' : f(rain.cur) + ' mm'}`, (rain.yd - rain.cur) / RS_TH.rain, 0.9);
    }
    out.sort((a, b) => b.w - a.w);
    const shipData = mc.size > 0 && my.size > 0;
    return { dir: sign > 0 ? 'up' : 'down', delta, pc, py, reasons: out.slice(0, RS_TH.top), every: out, all: out.length, shipData, rainData: !!(rain && Number.isFinite(rain.cur) && Number.isFinite(rain.yd)) };
  }

  // Lỗi do phiên đăng nhập hết hạn / chưa đăng nhập
  function isAuthError(e) {
    return /HTTP\s*(401|403)\b|đăng nhập|dang nhap|login|log in|logged|unauthor|expired|hết hạn|het han|session|phiên|\btoken\b|Unexpected token|not valid JSON/i.test(String((e && e.message) ?? e ?? ''));
  }

  /* ---- Cần thêm để đạt mục tiêu (ngưỡng xanh) ---- */
  const NEED_SRC = { pod: 'podN', podTT: 'podTTN', gtcTotal: 'gtcN', odrPct: 'odrPctN' };
  // m: chỉ số của một bưu cục; trả về số đơn cần giao thành công thêm để chạm ngưỡng xanh; 0 nếu đã đạt; null nếu không tính được
  function needMore(k, m, user) {
    const src = NEED_SRC[k], pair = m && src ? m[src] : null;
    if (!Array.isArray(pair) || !Number.isFinite(pair[0]) || !(pair[1] > 0)) return null;
    const t = Object.assign({}, ANM_TH[k], user && user[k]);
    if (!Number.isFinite(t.hi)) return null;
    return Math.max(0, Math.ceil(t.hi / 100 * pair[1] - pair[0] - 1e-9));
  }

  /* ---- Dự báo POD2400 cuối ngày: nền thống kê từ các ngày trước + khung gọi AI ---- */
  const FC_DAYS = 28, FC_MAXADJ = 5;
  const medianOf = a => { const s = [...a].sort((x, y) => x - y), n = s.length; return n ? (n % 2 ? s[(n - 1) / 2] : (s[n / 2 - 1] + s[n / 2]) / 2) : null; };
  const wdOf = dt => new Date(`${dt}T12:00:00`).getDay();
  // getAt(dt) = POD tại mốc giờ của ngày dt; getFin(dt) = POD cuối ngày dt (hoặc null)
  function forecastPairs(getAt, getFin, base, days = FC_DAYS) {
    const out = [];
    for (let i = 1; i <= days; i++) {
      const dt = addDays(base, -i), at = getAt(dt), fin = getFin(dt);
      if (Number.isFinite(at) && Number.isFinite(fin)) out.push({ date: dt, at, fin, wd: wdOf(dt) });
    }
    return out;
  }
  // Mức tăng POD từ mốc giờ đến cuối ngày (trung vị các ngày trước, ưu tiên cùng thứ nếu có ≥ 3 ngày)
  function forecastBase(podNow, pairs, wd) {
    if (!Number.isFinite(podNow)) return { ok: false, n: 0 };
    const all = (pairs || []).filter(p => Number.isFinite(p.at) && Number.isFinite(p.fin));
    if (all.length < 3) return { ok: false, n: all.length };
    const same = all.filter(p => p.wd === wd), use = same.length >= 3 ? same : all;
    const gains = use.map(p => p.fin - p.at), med = medianOf(gains);
    const mad = medianOf(gains.map(g => Math.abs(g - med)));
    const spread = Math.max(1, Math.round(1.4826 * mad * 2) / 2);
    const fc = Math.max(0, Math.min(100, podNow + med)), n = use.length;
    const conf = (n >= 14 && spread <= 2.5) ? 'Cao' : (n < 7 || spread > 4) ? 'Thấp' : 'Vừa';
    return { ok: true, n, fc, spread, conf, gain: med, sameWd: use === same };
  }
  const AI_PROMPT_DEFAULT = [
    'Bạn là trợ lý phân tích vận hành bưu cục chuyển phát tại Việt Nam. Trả lời bằng tiếng Việt.',
    'POD2400 = tỷ lệ đơn mới giao thành công trong 24 giờ. Mỗi bưu cục có "baseline": con số dự báo POD cuối ngày do chương trình tính từ lịch sử các mốc giờ trước (đã chính xác về mặt toán).',
    'Việc của bạn: xem tình trạng hôm nay (so hôm qua cùng mốc giờ: số DM đi làm, sản lượng, backlog 5 ngày, đơn chưa quét giao, giờ giao thành công đầu tiên, mưa) rồi quyết định có chỉnh baseline lên hoặc xuống không.',
    'Quy tắc: (1) KHÔNG tự tính lại con số từ đầu, chỉ chỉnh baseline; (2) mức chỉnh tối đa ±5 điểm, thường 0 đến ±2; (3) mỗi lý do phải nêu số cụ thể lấy từ dữ liệu được gửi; (4) thiếu dữ liệu thì đừng chỉnh và đừng bịa nguyên nhân; (5) mỗi bưu cục tối đa 3 lý do, mỗi lý do dưới 120 ký tự.',
    'Chỉ trả về MỘT khối JSON, không thêm chữ nào khác, đúng dạng: {"offices":[{"office":"<đúng tên đã gửi>","forecast":<số %, một chữ số thập phân>,"reasons":["...","..."]}]}'
  ].join('\n');
  const finite = v => (typeof v === 'number' && Number.isFinite(v)) ? Math.round(v * 10) / 10 : null;
  // items: [{ office, podNow, podYd, podWk, base:{fc,spread,conf,n}, backlog5, unscanned, scanRate, total, totalYd, dm, dmYd, firstOk, firstOkYd, finals, rain:{cur,yd} }]
  function buildAiPayload({ day, slot, items }) {
    return {
      date: day, slot,
      offices: items.map(i => ({
        office: i.office, pod_now: finite(i.podNow), pod_yesterday_same_slot: finite(i.podYd), pod_last_week_same_slot: finite(i.podWk),
        baseline: { forecast: finite(i.base.fc), spread: finite(i.base.spread), confidence: i.base.conf, days_used: i.base.n },
        backlog_5d: finite(i.backlog5), unscanned: finite(i.unscanned), scan_rate_pct: finite(i.scanRate),
        orders_total: finite(i.total), orders_total_yesterday: finite(i.totalYd), dm_working: finite(i.dm), dm_working_yesterday: finite(i.dmYd),
        first_success_time: Number.isFinite(i.firstOk) ? hmOf(i.firstOk) : null, first_success_time_yesterday: Number.isFinite(i.firstOkYd) ? hmOf(i.firstOkYd) : null,
        final_pod_last_days: (i.finals || []).map(finite),
        rain_mm_today: i.rain ? finite(i.rain.cur) : null, rain_mm_yesterday: i.rain ? finite(i.rain.yd) : null
      }))
    };
  }
  function aiRequestBody(model, system, userText, maxTokens) {
    return { model, max_tokens: maxTokens, system, messages: [{ role: 'user', content: userText }] };
  }
  // baseMap: { office: { fc, spread, conf } } ; trả về { office: { fc, low, high, conf, reasons, adj, clamped, fromBase } }
  function parseAiReply(text, baseMap) {
    let obj = null;
    const s = String(text ?? ''), a = s.indexOf('{'), b = s.lastIndexOf('}');
    if (a >= 0 && b > a) { try { obj = JSON.parse(s.slice(a, b + 1)); } catch { obj = null; } }
    const list = obj && Array.isArray(obj.offices) ? obj.offices : null;
    if (!list) throw new Error('AI trả về không đúng định dạng');
    const byName = new Map(); for (const o of list) if (o && typeof o.office === 'string') byName.set(o.office.trim(), o);
    const out = {};
    for (const [name, base] of Object.entries(baseMap)) {
      const o = byName.get(name), fcRaw = o ? Number(o.forecast) : NaN;
      if (!o || !Number.isFinite(fcRaw)) { out[name] = { fc: base.fc, low: Math.max(0, base.fc - base.spread), high: Math.min(100, base.fc + base.spread), conf: base.conf, reasons: [], adj: 0, clamped: false, fromBase: true }; continue; }
      const lo = Math.max(0, base.fc - FC_MAXADJ), hi = Math.min(100, base.fc + FC_MAXADJ), fc = Math.min(hi, Math.max(lo, fcRaw));
      const reasons = (Array.isArray(o.reasons) ? o.reasons : []).filter(x => typeof x === 'string' && x.trim()).slice(0, 3).map(x => x.trim().slice(0, 140));
      out[name] = { fc, low: Math.max(0, fc - base.spread), high: Math.min(100, fc + base.spread), conf: base.conf, reasons, adj: fc - base.fc, clamped: fc !== fcRaw, fromBase: false };
    }
    return out;
  }
  // Thông báo lỗi dễ hiểu khi gọi API AI
  function aiErrMsg(status, bodyText, netErr) {
    if (netErr) return 'Không kết nối được tới dịch vụ AI (mạng hoặc trang đang chặn kết nối ra ngoài)';
    const t = String(bodyText ?? '').toLowerCase();
    if (status === 401) return 'Khóa API không hợp lệ';
    if (status === 403) return 'Khóa API không có quyền sử dụng';
    if (status === 429) return 'Bị giới hạn tần suất, thử lại sau ít phút';
    if (/credit|balance|billing/.test(t)) return 'Tài khoản chưa có số dư (nạp tiền trong Console)';
    if (status === 404 || (status === 400 && /model/.test(t))) return 'Không tìm thấy mô hình AI, kiểm tra lại tên mô hình';
    if (status >= 500) return 'Dịch vụ AI đang lỗi, thử lại sau';
    return `Lỗi ${status || ''}`.trim();
  }

  /* ---- COD ship nộp tiền (tab ANM) ----
   * Mỗi dòng của API listCodPayment = một ship trong một ngày phát sinh (signDate).
   * Phải nộp = totalCharge, đã nộp = receivedCharge, chưa nộp = notSuccessReceivedCharge.
   * Chỉ dùng số tổng theo ship/ngày; không có mã vận đơn hay thông tin người nhận. */
  const COD_MIN = 1000, COD_LOOK = 7, COD_HALF = 0.5, COD_TODAY_ALERT_H = 23;
  // Ngày đang hiện: trước 12:00 là hôm qua (đang chốt), từ 12:00 là hôm nay (theo dõi luôn)
  function codView(now) {
    const h = now.getHours(), today = dateKey(now);
    return h < 12 ? { day: addDays(today, -1), isToday: false, track: true, today } : { day: today, isToday: true, track: true, today };
  }
  const codDateKey = s => { const m = String(s ?? '').match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/); return m ? `${m[3]}-${pad(m[2])}-${pad(m[1])}` : ''; };
  const codName = s => String(s ?? '').replace(/\s+/g, ' ').replace(/\s*\(\s*([^)]*?)\s*\)/, ' ($1)').trim();
  function codRow(r) {
    if (!r || typeof r !== 'object') return null;
    const date = codDateKey(r.signDate), id = String(r.dispatcherId ?? r.dispatcherCode ?? '').trim();
    if (!date || !id) return null;
    const due = nn(r.totalCharge) ?? nn(r.totalCodCahrge);
    if (due === null) return null;
    const paid = nn(r.receivedCharge) ?? 0, owe = nn(r.notSuccessReceivedCharge) ?? Math.max(0, due - paid);
    const pt = nn(r.paymentTime);
    return { id, name: codName(r.dispatcherName) || id, date, due, paid, owe, cod: nn(r.totalCodCahrge), n: nn(r.signCodCount), pt: pt !== null && pt > 1e11 ? pt : null, bad: nn(r.badDebtConfirmedAmount) || 0, site: String(r.signSiteName ?? '') };
  }
  // Gộp các trang: mỗi (ship, ngày) một dòng
  function codCollect(rawRows) {
    const m = new Map();
    for (const x of rawRows || []) { const r = codRow(x); if (r) m.set(`${r.id}|${r.date}`, r); }
    return [...m.values()];
  }
  // '' = đã nộp đủ (còn nợ dưới 1.000đ); r = nộp 0đ; o = nộp dưới 50%; y = nộp từ 50% nhưng còn nợ
  function codLevel(r) {
    if (!r || !(r.owe >= COD_MIN)) return '';
    if (!(r.paid >= 1)) return 'r';
    return r.due > 0 && r.paid / r.due < COD_HALF ? 'o' : 'y';
  }
  const codAge = (today, date) => Math.round((Date.parse(`${today}T00:00:00Z`) - Date.parse(`${date}T00:00:00Z`)) / 864e5);
  // Danh sách cảnh báo THEO SHIPPER: cộng dồn COD các ngày trong 7 ngày gần nhất (hôm nay chỉ tính từ 23:00),
  // chỉ báo shipper còn chưa thu đủ. 🔴 chưa nộp đồng nào trên tổng COD đã giao; 🟠 nộp dưới 50%; 🟡 nộp từ 50% nhưng còn nợ.
  // Shipper đã nộp đủ (kể cả đã nghỉ việc) không xuất hiện. Vàng: trước 12:00 chỉ tính khi nợ có từ trước hôm qua.
  function codAlerts(byOffice, now) {
    const D = dateKey(now), Y = addDays(D, -1), h = now.getHours(), from = addDays(D, -COD_LOOK), out = [], rank = { r: 0, o: 1, y: 2 };
    for (const [office, rows] of Object.entries(byOffice || {})) {
      const g = new Map();
      for (const r of rows || []) {
        if (!r || r.date > D || r.date < from || (r.date === D && h < COD_TODAY_ALERT_H)) continue;
        const k = r.id || r.name; let x = g.get(k);
        if (!x) g.set(k, x = { office, id: r.id, name: r.name, due: 0, paid: 0, owe: 0, cod: 0, n: 0, pt: NaN, bad: 0, date: '', days: 0 });
        x.due += r.due; x.paid += r.paid; x.owe += r.owe; x.cod += r.cod || 0; x.n += r.n || 0; x.bad += r.bad || 0;
        if (Number.isFinite(r.pt) && !(r.pt <= x.pt)) x.pt = r.pt;
        if (r.owe >= COD_MIN) { x.days++; if (!x.date || r.date < x.date) x.date = r.date; }
      }
      for (const x of g.values()) {
        const lv = codLevel(x);
        if (!lv) continue;
        if (lv === 'y' && h < 12 && !(x.date && x.date < Y)) continue;
        if (lv === 'y' && x.date === D) continue;   // chỉ còn nợ của hôm nay mà đã nộp từ 50%: chưa báo
        out.push(Object.assign(x, { lv, age: codAge(D, x.date) }));
      }
    }
    return out.sort((a, b) => rank[a.lv] - rank[b.lv] || b.owe - a.owe || a.office.localeCompare(b.office));
  }
  // Tổng theo bưu cục của đúng một ngày
  function codSummary(byOffice, offices, day) {
    const by = {}, T = { due: 0, paid: 0, owe: 0, cod: 0, zero: 0, ships: 0, loaded: 0 };
    for (const o of offices || []) {
      const rows = byOffice && byOffice[o];
      if (!rows) { by[o] = null; continue; }
      const rs = rows.filter(r => r.date === day), s = { due: 0, paid: 0, owe: 0, cod: 0, zero: 0, ships: rs.length };
      for (const r of rs) { s.due += r.due; s.paid += r.paid; s.owe += r.owe; s.cod += r.cod || 0; if (codLevel(r) === 'r') s.zero++; }
      s.pct = ratio(s.paid, s.due); by[o] = s; T.loaded++;
      for (const k of ['due', 'paid', 'owe', 'cod', 'zero', 'ships']) T[k] += s[k];
    }
    T.pct = ratio(T.paid, T.due);
    return { by, tot: T };
  }
  const grp = n => String(Math.round(Math.abs(n))).replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  function fmtVnd(v) {
    if (!Number.isFinite(v)) return '—';
    const a = Math.abs(v), r = Math.round(a * 100) / 100, ip = Math.floor(r), fr = Math.round((r - ip) * 100);   // đủ chữ số, không làm tròn; có số lẻ thì hiện số lẻ
    return `${v < 0 ? '-' : ''}${String(ip).replace(/\B(?=(\d{3})+(?!\d))/g, '.')}${fr ? ',' + String(fr).padStart(2, '0').replace(/0$/, '') : ''}đ`;
  }
  const fmtPct1 = v => Number.isFinite(v) ? `${v.toFixed(1).replace('.', ',')}%` : '—';
  // Giờ nộp gần nhất: “22:17 hôm qua”
  function codLastPay(pt, now) {
    if (!Number.isFinite(pt)) return 'chưa nộp lần nào';
    const t = new Date(pt), k = dateKey(t), D = dateKey(now), hm = `${pad(t.getHours())}:${pad(t.getMinutes())}`;
    return `${hm} ${k === D ? 'hôm nay' : k === addDays(D, -1) ? 'hôm qua' : fmtDM(k)}`;
  }
  const COD_ICON = { r: '🔴', o: '🟠', y: '🟡' };
  function codCopyText(alerts, now) {
    const hm = `${pad(now.getHours())}:${pad(now.getMinutes())} ${fmtDM(dateKey(now))}`;
    if (!alerts.length) return `COD ship: không có ship nào cần cảnh báo (cập nhật ${hm})`;
    return [`⚠ COD ship chưa nộp · cập nhật ${hm}`].concat(alerts.map(a =>
      `${COD_ICON[a.lv]} ${a.office} · ${a.name} · nợ ${fmtVnd(a.owe)} / ${fmtVnd(a.due)} (${a.lv === 'r' ? 'chưa nộp đồng nào' : 'đã nộp ' + fmtVnd(a.paid) + ' = ' + fmtPct1(ratio(a.paid, a.due))}) · chưa thu đủ, nợ từ ${fmtDM(a.date)}`)).join('\n');
  }

  // Ship nộp 0đ trong ngày day của các bưu cục offices (xếp theo tiền chưa nộp giảm dần)
  function codZeroList(byOffice, offices, day) {
    const out = [];
    for (const o of offices || []) for (const r of (byOffice && byOffice[o]) || []) if (r && r.date === day && codLevel(r) === 'r') out.push({ office: o, name: r.name, owe: r.owe, due: r.due });
    return out.sort((a, b) => b.owe - a.owe || a.office.localeCompare(b.office) || String(a.name).localeCompare(String(b.name)));
  }
  function codZeroText(list, label, day, now) {
    const hm = `${pad(now.getHours())}:${pad(now.getMinutes())} ${fmtDM(dateKey(now))}`;
    if (!list.length) return `COD ship nộp 0đ · ${label} · ${fmtDM(day)}: không có ship nào (cập nhật ${hm})`;
    return [`🔴 COD ship chưa nộp đồng nào · ${label} · ngày ${fmtDM(day)} · ${list.length} ship · tổng ${fmtVnd(list.reduce((t, x) => t + x.owe, 0))} · cập nhật ${hm}`].concat(list.map(x => `${x.office === label ? '' : x.office + ' · '}${x.name} · chưa nộp ${fmtVnd(x.owe)}`)).join('\n');
  }

  // ----- Tự dò bưu cục quản lý (hàm thuần) -----
  const DISC_SKIP_TYPE = ['DC', 'KA'];   // trung tâm trung chuyển (đơn gộp của các bưu cục con) và hàng cargo: không phải bưu cục giao hàng
  const discName = x => { const v = x && x.nameLocale; const t = typeof v === 'string' ? v : (v && typeof v === 'object' ? (Object.values(v).find(q => typeof q === 'string') || '') : (x && x.name) || ''); return String(t || '').trim().replace(/\s+/g, ' '); };
  // list: danh sách bưu cục hệ thống trả về; known: Set tên (đã bỏ dấu) đã có; hidden: Set tên người dùng đã xóa
  function discPick(list, known, hidden) {
    const out = { cand: [], closed: 0, skipType: [], removed: [], have: [], bad: 0, dup: 0, total: 0 };
    const seen = new Set(), seenName = new Set();
    for (const x of Array.isArray(list) ? list : []) {
      if (!x || typeof x !== 'object') { out.bad++; continue; }
      out.total++;
      const name = discName(x), id = x.id, code = x.code;
      if (!name || !id || code === undefined || code === null || code === '') { out.bad++; continue; }
      const key = String(id);
      if (seen.has(key)) { out.dup++; continue; }
      seen.add(key);
      if (x.status !== 'ENABLE') { out.closed++; continue; }
      if (seenName.has(name)) { out.dup++; continue; }   // tên là duy nhất; trùng tên thì chỉ lấy bưu cục đầu tiên đang mở
      seenName.add(name);
      if (DISC_SKIP_TYPE.includes(String(x.type || '').toUpperCase()) || String(x.logisticsType || '').toUpperCase() === 'CARGO') { out.skipType.push(name); continue; }
      if (known.has(plainOf(name))) { out.have.push(name); continue; }
      if (hidden.has(name) || hidden.has(plainOf(name))) { out.removed.push(name); continue; }
      out.cand.push({ name, id, code: String(code), type: String(x.type || '') });
    }
    return out;
  }
  const plainOf = s => String(s ?? '').normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/đ/g, 'd').replace(/Đ/g, 'D').replace(/\s+/g, ' ').trim().toUpperCase();
  // Lấy ownerSiteId từ nội dung lệnh pageQuery của trang (chuỗi JSON) – null nếu không có
  function discOwnerOf(body) {
    try { const o = typeof body === 'string' ? JSON.parse(body) : body; const v = o && o.ownerSiteId; const n = typeof v === 'string' && v.trim() ? Number(v) : v; return Number.isFinite(n) && n > 0 ? n : null; } catch { return null; }
  }
  function discSummary(pk, added, empty, failed, left) {
    const parts = [];
    parts.push(added.length ? `thêm ${added.length} bưu cục (${added.join(', ')})` : 'không có bưu cục mới');
    if (empty.length) parts.push(`bỏ qua ${empty.length} bưu cục rỗng (${empty.join(', ')})`);
    if (pk.closed) parts.push(`${pk.closed} đã đóng`);
    if (pk.skipType.length) parts.push(`${pk.skipType.length} trung tâm/cargo (${pk.skipType.join(', ')})`);
    if (pk.removed.length) parts.push(`${pk.removed.length} đã xóa trước đó nên không thêm lại`);
    if (failed.length) parts.push(`${failed.length} chưa kiểm tra được (${failed.join(', ')})`);
    if (left) parts.push(`còn ${left} bưu cục chưa kiểm tra, bấm tải lần nữa`);
    return `Dò bưu cục: ${parts.join('; ')}.`;
  }

  // </pure>

  const $ = (s, root = document) => root.querySelector(s);
  const norm = el => (el?.innerText || el?.textContent || '').replace(/\s+/g, ' ').trim().normalize('NFC');
  const visible = el => !!(el && (el.offsetWidth || el.offsetHeight || el.getClientRects().length));
  const sleep = ms => new Promise(r => setTimeout(r, ms));

  function log(msg, bad = false) {
    if (bad) state.lastErr = msg;
    const el = $('#b8s-log');
    if (el) {
      const line = document.createElement('div');
      if (bad) line.className = 'bad';
      line.textContent = `[${new Date().toLocaleTimeString('vi-VN')}] ${msg}`;
      el.prepend(line);
      while (el.children.length > 100) el.lastChild.remove();
    }
    const st = $('#b8s-status');
    if (st) { st.textContent = msg; st.className = bad ? 'bad' : ''; }
  }

  /* ------------------------------------------------------------------ *
   * Ghi nhận cấu trúc API (chỉ tên khóa) để phục vụ chẩn đoán
   * ------------------------------------------------------------------ */
  function listOf(payload) {
    const d = payload?.data ?? payload;
    if (Array.isArray(d)) return d;
    if (d && typeof d === 'object') for (const k of ['list', 'records', 'rows', 'items', 'resultList', 'dataList']) if (Array.isArray(d[k])) return d[k];
    return null;
  }

  function record(url, json) {
    try {
      const u = new URL(url, location.href);
      if (u.origin !== location.origin || !/\/(v9-|sea-)/.test(u.pathname)) return;
      const list = listOf(json), d = json?.data, prev = state.endpoints.get(u.pathname);
      state.endpoints.set(u.pathname, {
        count: (prev?.count || 0) + 1,
        topKeys: json && typeof json === 'object' ? Object.keys(json).slice(0, 12) : [],
        dataKeys: d && typeof d === 'object' && !Array.isArray(d) ? Object.keys(d).slice(0, 25) : undefined,
        listLength: list ? list.length : undefined,
        itemKeys: list && list[0] && typeof list[0] === 'object' ? Object.keys(list[0]).slice(0, 40) : undefined
      });
    } catch { /* bỏ qua */ }
  }

  // Bắt tiêu đề xác thực (token…) của các lệnh API trang gửi. Thư viện của trang có thể ghi đè hàm của trình duyệt,
  // nên kiểm tra định kỳ và gắn lại lên trên cùng nếu hàm của tool đã bị thay.
  function ensureHdrHook() {
    try {
      const P = XMLHttpRequest.prototype;
      if (P.setRequestHeader && P.setRequestHeader.__b8) return;
      const cur = P.setRequestHeader;
      const ours = function (k, v) {
        try { if (/^(token|lang-type|langType|TIME-ZONE|no-offset-timeZone|data-masking-N|Accept)$/i.test(k)) state.apiHdr[k] = v; } catch { /* bỏ qua */ }
        return cur.apply(this, arguments);
      };
      ours.__b8 = true;
      P.setRequestHeader = ours;
    } catch { /* bỏ qua */ }
  }

  // Trang tự gửi lệnh danh sách bưu cục kèm ownerSiteId của chính tài khoản đang dùng: ghi lại để tự dò bưu cục
  function captureOwner(body) {
    const n = discOwnerOf(body);
    if (n && n !== state.ownerSiteId) { state.ownerSiteId = n; try { localStorage.setItem('b8s_owner', String(n)); } catch { /* bỏ qua */ } }
  }
  function hookNetwork() {
    const nativeFetch = window.fetch;
    if (typeof nativeFetch === 'function') {
      window.fetch = function (...args) {
        const p = nativeFetch.apply(this, args);
        try {
          const a = args[0];
          const url = typeof a === 'string' ? a : (a?.url || String(a || ''));
          if (/\/sea-man\/node\/pageQuery/.test(url) && args[1] && typeof args[1].body === 'string') captureOwner(args[1].body);
          p.then(res => res.clone().json().then(j => record(url, j)).catch(() => {})).catch(() => {});
        } catch { /* bỏ qua */ }
        return p;
      };
    }
    ensureHdrHook(); setInterval(ensureHdrHook, 1000);
    const xo = XMLHttpRequest.prototype.open, xs = XMLHttpRequest.prototype.send;
    XMLHttpRequest.prototype.open = function (method, url, ...rest) { this.__b8surl = String(url); return xo.call(this, method, url, ...rest); };
    XMLHttpRequest.prototype.send = function (...args) {
      try { if (this.__b8surl && /\/sea-man\/node\/pageQuery/.test(this.__b8surl)) captureOwner(args[0]); } catch { /* bỏ qua */ }
      this.addEventListener('load', () => {
        try { if (this.__b8surl && (this.responseType === '' || this.responseType === 'text')) record(this.__b8surl, JSON.parse(this.responseText)); } catch { /* không phải JSON */ }
      });
      return xs.apply(this, args);
    };
  }

  /* ------------------------------------------------------------------ *
   * Đọc dữ liệu từ trang
   * ------------------------------------------------------------------ */
  // Ô “Bưu cục” trên trang là ô nhập/chọn, nên tên nằm trong value chứ không nằm trong chữ hiển thị
  function findOfficeInput() {
    const labels = [...document.querySelectorAll('label, span, div')]
      .filter(el => visible(el) && !el.closest('#b8s-app') && /^Bưu cục\s*[:：]?$/i.test(norm(el)));
    for (const el of labels) {
      let p = el;
      for (let i = 0; i < 4 && p; i++) {
        p = p.parentElement;
        const inp = p && [...p.querySelectorAll('input')].find(x => !/checkbox|radio|hidden/.test(x.type) && !x.closest('#b8s-app') && (visible(x) || x.closest('.ant-select, .el-select, .ivu-select')));
        if (inp) return inp;
      }
    }
    return null;
  }

  // Lưu ý: ô gõ của ant-select có class “ant-select-search__field” (chứa chữ “select”) nên phải loại chính nó khi tìm khung bao ngoài
  const selectWrap = inp => inp && (inp.closest('.ant-select, .el-select, .ivu-select') || (inp.parentElement && inp.parentElement.closest('[class*="select"]')) || inp.parentElement);
  function currentOfficeText(inp) {
    if (!inp) return '';
    const w = selectWrap(inp);
    // Tên đã chọn nằm ở thẻ hiển thị giá trị, không nằm trong ô gõ (ô gõ luôn trống sau khi chọn)
    const sv = w && w.querySelector && w.querySelector('.ant-select-selection-selected-value, .el-select__tags-text, .ivu-select-selected-value');
    const sel = sv && ((sv.getAttribute('title') || sv.textContent || '').trim());
    if (sel) return sel.normalize('NFC');
    let v = (inp.value || '').trim();
    if (!v) {
      v = norm(w);
      const ph = (inp.placeholder || '').trim();
      if (ph && v === ph) v = '';
      if (/^xin chọn/i.test(v)) v = '';
    }
    return v;
  }

  function readOfficeInfo() {
    const text = document.body?.innerText || '';
    const inp = findOfficeInput();
    let office = currentOfficeText(inp);
    if (!office) office = parseOfficeText(text);
    return { office, province: parseProvince(text), text };
  }

  function rememberOffice(name) {
    state.recent = [name, ...state.recent.filter(x => plain(x) !== plain(name))].slice(0, 20);
    savePrefs(); renderOfficeList();
  }

  function applyOfficeInfo({ office, province }) {
    if (province) state.province = province;
    if (!office || office === state.office) return;
    if (state.office) {
      state.emp = null; state.orders = null; state.ov = null; state.pending = null;
      log(`Đã đổi bưu cục: ${office}. Dữ liệu cũ được xóa.`);
    }
    state.office = office;
    rememberOffice(office);
  }

  // Số liệu trang tổng quan: chỉ để đối chiếu với số tự tính
  function readOverview() {
    const info = readOfficeInfo();
    applyOfficeInfo(info);
    const text = info.text;
    const get = re => { const m = text.match(re); return m ? parseNum(m[1]) : NaN; };
    const need = get(/Tổng đơn cần phát hàng\s*([\d,.]+)/i);
    const unscanned = get(/Đã nhập kho nhưng chưa quét phát\s*([\d,.]+)/i);
    if (Number.isFinite(need) || Number.isFinite(unscanned)) state.ov = { need, unscanned, at: new Date() };
  }

  const FIXED = '.el-table__fixed, .el-table__fixed-right, .vxe-table--fixed-left-wrapper, .vxe-table--fixed-right-wrapper, .fixed-left--wrapper, .fixed-right--wrapper, .ant-table-fixed-left, .ant-table-fixed-right';
  const TABLE_SEL = '.el-table, .ant-table, .vxe-table, table, [role="grid"], [role="table"]';

  function findEmployeeRoot() {
    let best = null;
    for (const r of document.querySelectorAll(TABLE_SEL)) {
      if (!visible(r) || r.closest(FIXED)) continue;
      const t = norm(r);
      if (t.includes('Nhân viên giao nhận') && t.includes('Đơn cần phát') && t.includes('Đã quét phát')) {
        if (!best || best.contains(r)) best = r;
      }
    }
    return best;
  }

  function readTable(root) {
    const headRows = [...root.querySelectorAll('thead tr')].filter(tr => !tr.closest(FIXED));
    let headCells = [];
    for (const tr of headRows) {
      const cells = [...tr.children].filter(c => /^(TH|TD)$/.test(c.tagName) && !c.classList.contains('gutter'));
      if (cells.length > headCells.length) headCells = cells;
    }
    if (!headCells.length) headCells = [...root.querySelectorAll('[role="columnheader"]')];
    const rawHeaders = headCells.map(norm);

    const rows = [];
    const trs = [...root.querySelectorAll('tbody tr')].filter(tr =>
      !tr.closest(FIXED) && !tr.classList.contains('ant-table-measure-row') && visible(tr));
    for (const tr of trs) {
      const cells = [...tr.children].filter(c => /^(TD|TH)$/.test(c.tagName)).map(norm);
      if (cells.length < 3 || !cells[0]) continue;
      if (/^(Tổng|合计|Không có dữ liệu|暂无数据)/i.test(cells[0])) continue;
      rows.push(cells);
    }
    const width = Math.max(rawHeaders.length, ...rows.map(r => r.length), 0);
    const headers = uniqueHeaders(rawHeaders, width);
    rows.forEach(r => { while (r.length < width) r.push(''); });
    return { headers, rows };
  }

  function findOrderRoot() {
    const emp = findEmployeeRoot();
    for (const r of document.querySelectorAll(TABLE_SEL)) {
      if (!visible(r) || r.closest(FIXED) || r === emp || (emp && emp.contains(r))) continue;
      const t = readTable(r);
      if (t.rows.length && waybillCol(t.headers, t.rows) >= 0) return r; // phần tử ngoài cùng (có cả tiêu đề)
    }
    return null;
  }

  /* ------------------------------------------------------------------ *
   * Gom dữ liệu qua nhiều trang, tự mở tab và quay lại tab cũ
   * ------------------------------------------------------------------ */
  const isDisabled = el => !el || el.disabled || el.hasAttribute('disabled') ||
    el.getAttribute('aria-disabled') === 'true' || /\b(disabled|ant-pagination-disabled)\b/.test(el.className || '');

  function findPager(root) {
    const rr = root.getBoundingClientRect();
    const list = [...document.querySelectorAll('.el-pagination, .ant-pagination')].filter(visible);
    const below = list.filter(p => p.getBoundingClientRect().top >= rr.top);
    const pool = below.length ? below : list;
    pool.sort((a, b) => Math.abs(a.getBoundingClientRect().top - rr.bottom) - Math.abs(b.getBoundingClientRect().top - rr.bottom));
    return pool[0] || null;
  }

  // Chữ ký nhẹ: số dòng + nội dung dòng đầu/cuối (không phân tích cả bảng)
  function tableSig(root) {
    const trs = [...root.querySelectorAll('tbody tr')].filter(tr => !tr.closest(FIXED));
    if (!trs.length) return '0';
    const t = tr => (tr.textContent || '').replace(/\s+/g, ' ').slice(0, 240);
    return trs.length + '|' + t(trs[0]) + '|' + t(trs[trs.length - 1]);
  }

  const isLoading = () => [...document.querySelectorAll('.el-loading-mask, .ant-spin-spinning')].some(visible);

  async function waitForChange(sig, finder) {
    const t0 = Date.now();
    while (Date.now() - t0 < 8000) {
      await sleep(100);
      const root = finder();
      if (root && !isLoading() && tableSig(root) !== sig) { await sleep(80); return true; }
    }
    return false;
  }

  async function goFirstPage(finder) {
    const root = finder();
    const pg = root && findPager(root);
    const first = pg && pg.querySelector('.el-pager li.number, .ant-pagination-item-1');
    if (!first || /\b(active|ant-pagination-item-active)\b/.test(first.className)) return;
    const sig = tableSig(root);
    first.click();
    await waitForChange(sig, finder);
  }

  const EMP_TAB = 'Tiến độ nhân viên';
  const isEmpTabText = t => t.normalize('NFC') === EMP_TAB;
  const isOrderTabText = t => t.length <= 40 && /^chi tiết phát/i.test(t.normalize('NFC'));

  function findTab(pred) {
    const hits = [...document.querySelectorAll('.el-tabs__item, [role="tab"], .ant-tabs-tab, li, a, span, div, button')]
      .filter(el => visible(el) && !el.closest('#b8s-app') && norm(el).length <= 60 && pred(norm(el)));
    return hits.find(el => el.matches('.el-tabs__item, [role="tab"], .ant-tabs-tab')) || hits[hits.length - 1] || null;
  }

  async function openTab(finder, pred, label) {
    if (finder()) return { switched: false };
    const target = findTab(pred);
    if (!target) { log(`Không tìm thấy tab “${label}” trên trang.`, true); return null; }
    const group = target.closest('.el-tabs__nav, .ant-tabs-nav, [role="tablist"]');
    const prev = group?.querySelector('.is-active, [aria-selected="true"], .ant-tabs-tab-active') || null;
    log(`Tự chuyển sang tab “${label}”…`);
    target.click();
    const t0 = Date.now();
    while (Date.now() - t0 < 12000) {
      await sleep(150);
      const root = finder();
      if (root && !isLoading() && root.querySelector('tbody tr')) { await sleep(150); return { switched: true, prev }; }
    }
    log(`Đã chuyển tab “${label}” nhưng chưa có dữ liệu. Nếu tab cần bấm Tìm kiếm/Truy vấn, hãy bấm trước rồi chạy lại.`, true);
    return { switched: true, prev };
  }

  async function restoreTab(h) {
    if (h && h.switched && h.prev && document.contains(h.prev)) { h.prev.click(); await sleep(200); }
  }

  // Chọn số dòng/trang lớn nhất để cần ít lần bấm trang nhất
  async function setBigPage(finder) {
    try {
      const root = finder();
      const pg = root && findPager(root);
      const box = pg && pg.querySelector('.el-pagination__sizes .el-input, .el-pagination__sizes, .ant-pagination-options-size-changer');
      if (!box) return;
      const inp = box.querySelector('input');
      const cur = inp ? parseNum(inp.value) : NaN;
      const sig = tableSig(root);
      (box.querySelector('.el-input, .ant-select-selector') || box).click();
      await sleep(250);
      const items = [...document.querySelectorAll('.el-select-dropdown__item, .ant-select-item-option')]
        .filter(el => visible(el) && !el.closest('#b8s-app'));
      let best = null, bestN = Number.isFinite(cur) ? cur : 0;
      for (const el of items) { const n = parseNum(norm(el)); if (Number.isFinite(n) && n > bestN) { best = el; bestN = n; } }
      if (!best) { inp?.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', keyCode: 27, bubbles: true })); return; }
      best.click();
      await waitForChange(sig, finder);
      log(`Đã chọn ${bestN} dòng/trang để tải nhanh hơn.`);
    } catch { /* không đổi được thì đọc từng trang như cũ */ }
  }

  async function collectPages({ finder, pred, label, maxPages = 300, dedupe = true }) {
    let handle = null;
    try {
      handle = await openTab(finder, pred, label);
      if (!handle || !finder()) return null;
      await goFirstPage(finder);
      await setBigPage(finder);
      const seen = new Map(), raw = [];
      let headers = [], pagerTotal = 0;
      for (let page = 1; page <= maxPages; page++) {
        const root = finder();
        if (!root) { log('Mất bảng giữa chừng, dừng.', true); break; }
        const t = readTable(root);
        headers = t.headers;
        if (dedupe) t.rows.forEach(r => seen.set(JSON.stringify(r), r)); else raw.push(...t.rows);
        log(`${label}: đã đọc trang ${page} (${dedupe ? seen.size : raw.length} dòng)…`);
        const pg = findPager(root);
        const totalEl = pg && pg.querySelector('.el-pagination__total, .ant-pagination-total-text');
        if (totalEl) { const tn = parseNum(norm(totalEl)); if (Number.isFinite(tn)) pagerTotal = tn; }
        const next = pg && pg.querySelector('.btn-next, .ant-pagination-next');
        if (!next || isDisabled(next)) break;
        const sig = tableSig(root);
        next.click();
        if (!(await waitForChange(sig, finder))) { log('Trang kế tiếp không đổi nội dung, dừng.', true); break; }
      }
      await goFirstPage(finder);
      const rows = dedupe ? [...seen.values()] : raw;
      return rows.length ? { headers, rows, pagerTotal } : null;
    } finally {
      await restoreTab(handle);
    }
  }

  async function runBusy(fn) {
    if (state.busy) return false;
    state.busy = true; setBusy(true);
    try { return await fn(); }
    catch (e) { log(`Lỗi: ${e.message}`, true); return false; }
    finally { state.busy = false; setBusy(false); renderUI(); }
  }

  async function collectEmployees() {
    const res = await collectPages({ finder: findEmployeeRoot, pred: isEmpTabText, label: EMP_TAB });
    if (!res) { log('Không đọc được bảng Tiến độ nhân viên.', true); return false; }
    state.emp = { headers: res.headers, rows: res.rows, at: new Date() };
    log(`Tiến độ nhân viên: ${res.rows.length} dòng.`);
    return true;
  }

  const PII = /người nhận|nhận hàng|khách nhận|sđt|số đt|điện thoại|phone|\btel\b|địa chỉ|address|收件|电话/i;

  async function collectOrders() {
    const res = await collectPages({ finder: findOrderRoot, pred: isOrderTabText, label: 'Chi tiết phát kiện', maxPages: 500, dedupe: false });
    if (!res) { log('Không đọc được bảng Chi tiết phát kiện. Mở Cài đặt > Chẩn đoán API và gửi kết quả để chỉnh.', true); return false; }
    const wb = waybillCol(res.headers, res.rows);
    if (wb < 0) { log('Không nhận diện được cột mã vận đơn.', true); return false; }

    // Mỗi vận đơn chỉ tính một lần (giữ dòng đầy đủ nhất)
    const filled = r => r.filter(v => String(v ?? '').trim() !== '').length;
    const best = new Map();
    let noWb = 0;
    for (const r of res.rows) {
      const k = String(r[wb] ?? '').trim().toUpperCase();
      if (!k) { noWb++; continue; }
      const cur = best.get(k);
      if (!cur || filled(r) > filled(cur)) best.set(k, r);
    }
    const uniq = [...best.values()];
    log(`Đã đọc ${res.rows.length} dòng → ${uniq.length} vận đơn duy nhất${noWb ? ` (${noWb} dòng không có mã)` : ''}.`);
    if (res.pagerTotal) {
      if (res.pagerTotal === uniq.length) log(`Khớp với tổng ${fmt(res.pagerTotal)} của bộ phân trang.`);
      else log(`Chú ý: bộ phân trang báo ${fmt(res.pagerTotal)} nhưng đọc được ${fmt(uniq.length)} vận đơn.`, true);
    }

    // Mặc định bỏ cột thông tin người nhận ngay khi đọc
    const idx = res.headers.map((h, i) => (state.includePII || i === wb || !PII.test(h)) ? i : -1).filter(i => i >= 0);
    state.orders = {
      headers: idx.map(i => res.headers[i]),
      rows: uniq.map(r => idx.map(i => r[i])),
      wbCol: idx.indexOf(wb),
      at: new Date()
    };
    return true;
  }

  // Đọc tổng quan -> tiến độ nhân viên -> chi tiết phát kiện
  async function doUpdate(expect) {
    if (expect) applyOfficeInfo({ office: expect, province: '' });
    readOverview();
    const startOffice = expect || state.office || '';
    state.lastSave = null;
    let e = false, o = false, viaApi = false;
    // Tải nhanh bằng API (một lần gọi thay vì đọc từng trang 50 dòng); lỗi thì tự quay về cách đọc bảng
    const apiOffice = expect || readOfficeInfo().office || state.office || '';
    if (state.fast && apiOffice) {
      try {
        const r = await apiFetchOffice(apiOffice);
        const now = new Date();
        state.emp = { headers: r.t.emp.headers, rows: r.t.emp.rows, at: now };
        state.orders = { headers: r.t.orders.headers, rows: r.t.orders.rows, wbCol: r.t.orders.wbCol, at: now };
        e = o = viaApi = true;
        log(`⚡ Tải nhanh bằng API: ${fmt(r.t.orders.rows.length)} đơn, ${fmt(r.t.emp.rows.length)} nhân viên (${apiOffice}).`);
      } catch (err) { log(`Tải nhanh bằng API lỗi: ${err.message}. Chuyển sang đọc từng trang.`, true); }
    }
    if (!viaApi) { e = await collectEmployees(); o = await collectOrders(); }
    state.lastRefresh = new Date();
    if (o && e) log(viaApi ? 'Hoàn tất cập nhật (API).' : 'Hoàn tất cập nhật.');
    state.lastSave = o ? await trySaveHistory(expect, startOffice) : { saved: false, reason: 'Không đọc được bảng Chi tiết phát kiện.' };
    return o || e;
  }
  const updateAll = () => runBusy(doUpdate);

  /* ------------------------------------------------------------------ *
   * Chọn bưu cục ngay trên tool (điều khiển ô “Bưu cục” và nút “Tra cứu” của trang)
   * ------------------------------------------------------------------ */
  const OPT_SEL = '.el-select-dropdown__item, .ant-select-item-option, .el-cascader-node, .el-tree-node__content, li[role="option"], .el-autocomplete-suggestion li, .ivu-select-item, .vue-treeselect__option';
  const getOptions = () => [...document.querySelectorAll(OPT_SEL)]
    .filter(el => visible(el) && !el.closest('#b8s-app') && !/\bis-disabled\b|\bdisabled\b/.test(el.className || ''));

  // Nút “Tra cứu” có thể là button, a, div hoặc span: lấy phần tử có chữ đúng “Tra cứu” nằm cùng hàng với ô Bưu cục (ưu tiên bên phải)
  function searchCandidates(inp) {
    const w = (selectWrap(inp) || inp).getBoundingClientRect();
    const cy = (w.top + w.bottom) / 2;
    const out = [];
    for (const el of document.querySelectorAll('button, a, span, div, li, [role="button"], input[type="button"], input[type="submit"]')) {
      if (el.closest('#b8s-app') || !visible(el)) continue;
      const t = plain(el.tagName === 'INPUT' ? el.value : norm(el));
      if (t !== 'tra cuu') continue;
      const b = el.getBoundingClientRect();
      const dy = Math.abs((b.top + b.bottom) / 2 - cy);
      const dx = b.left >= w.right - 5 ? b.left - w.right : 3000 + Math.abs(b.left - w.right);
      out.push({ el, dy, dx, score: dy * 3 + dx * 0.1 });
    }
    return out.sort((x, y) => x.score - y.score);
  }
  function findSearchBtn(inp) {
    const c = searchCandidates(inp).find(x => x.dy <= 90);
    return c ? c.el : null;
  }

  function closeDropdown(inp) {
    try {
      inp.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', keyCode: 27, bubbles: true }));
      inp.blur();
    } catch { /* bỏ qua */ }
  }

  const fireMouse = (el, types) => types.forEach(t => el.dispatchEvent(new MouseEvent(t, { bubbles: true, cancelable: true, view: window })));

  // Các dòng của danh sách thả xuống: phần tử lá, hiện ra ngay bên dưới ô Bưu cục
  function dropdownItems(inp) {
    const r = inp.getBoundingClientRect(), w = selectWrap(inp)?.getBoundingClientRect() || r;
    const out = [];
    for (const el of document.querySelectorAll('body *')) {
      if (el.childElementCount || el.closest('#b8s-app') || !visible(el)) continue;
      const t = norm(el);
      if (!t || t.length > 80) continue;
      const b = el.getBoundingClientRect();
      if (b.top < Math.max(r.bottom, w.bottom) - 6 || b.top > r.bottom + 700) continue;
      if (b.left > w.right + 40 || b.right < w.left - 40) continue;
      out.push(el);
    }
    return out;
  }

  function scrollParentOf(el) {
    for (let p = el.parentElement; p && p !== document.body; p = p.parentElement) {
      const cs = getComputedStyle(p);
      if (/(auto|scroll)/.test(cs.overflowY) && p.scrollHeight > p.clientHeight + 8) return p;
    }
    return null;
  }

  async function selectOffice(name) {
    const target = plain(name);
    if (!target) { log('Hãy chọn một bưu cục.', true); return false; }
    const inp = findOfficeInput();
    if (!inp) { log('Không thấy ô “Bưu cục” trên trang. Hãy mở tab “Tiến độ giao hàng”.', true); return false; }
    if (plain(currentOfficeText(inp)) === target) { log('Đang ở đúng bưu cục này, chỉ cập nhật dữ liệu.'); return true; }

    // 1) mở danh sách như khi bạn bấm vào ô
    const opener = inp.closest('.ant-select-selector, .el-input, .el-select, .ivu-select') || selectWrap(inp) || inp;
    try { inp.focus(); } catch { /* bỏ qua */ }
    fireMouse(opener, ['mousedown', 'mouseup', 'click']);
    await sleep(500);

    // 2) gõ tên để lọc (ô tìm kiếm trong danh sách)
    if (!inp.readOnly) {
      try {
        Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(inp, name);
        inp.dispatchEvent(new Event('input', { bubbles: true }));
      } catch { /* bỏ qua */ }
    }
    log(`Đang tìm bưu cục “${name}” trong danh sách…`);

    // 3) tìm dòng đúng tên; nếu danh sách dài thì cuộn dần
    let pick = null, scroller = null, lastTop = -1, still = 0;
    const t0 = Date.now();
    while (Date.now() - t0 < 15000) {
      await sleep(350);
      const items = dropdownItems(inp);
      if (!items.length) { fireMouse(opener, ['mousedown', 'mouseup', 'click']); continue; }
      const r = pickOption(items.map(norm), name);
      if (r.index >= 0) { pick = items[r.index]; break; }
      if (r.ambiguous.length && Date.now() - t0 > 4000) {
        log(`Có nhiều bưu cục khớp: ${r.ambiguous.slice(0, 5).join(' | ')}. Hãy chọn tên đầy đủ hơn.`, true);
        closeDropdown(inp);
        return false;
      }
      scroller = scroller || scrollParentOf(items[0]);
      if (scroller) {
        scroller.scrollTop += Math.max(120, scroller.clientHeight * 0.8);
        still = scroller.scrollTop === lastTop ? still + 1 : 0;
        lastTop = scroller.scrollTop;
        if (still >= 3) break; // đã cuộn tới cuối
      }
    }
    if (!pick) {
      log(`Không tìm thấy bưu cục “${name}” trong danh sách của trang (có thể bạn không được cấp quyền bưu cục này).`, true);
      closeDropdown(inp);
      return false;
    }

    pick.scrollIntoView?.({ block: 'nearest' });
    fireMouse(pick, ['mouseover', 'mousedown', 'mouseup', 'click']);
    await sleep(700);
    if (!plain(currentOfficeText(inp)).includes(target)) log('Ô Bưu cục chưa hiển thị đúng tên đã chọn; vẫn tiếp tục tra cứu.', true);

    // 4) trang tự tải số liệu ngay khi chọn bưu cục (không cần bấm Tra cứu): chờ tải xong và ổn định
    await sleep(1200);
    let quiet = 0;
    const t1 = Date.now();
    while (Date.now() - t1 < 30000 && quiet < 3) {
      quiet = isLoading() ? 0 : quiet + 1;
      await sleep(300);
    }
    await sleep(600);
    log(`Đã chuyển sang bưu cục: ${currentOfficeText(inp)}`);
    return true;
  }

  // Tab tổng quan: nếu bật tải nhanh thì không cần điều khiển ô Bưu cục của trang, gọi API thẳng cho bưu cục được chọn
  async function apiOverview(name) {
    const r = await apiFetchOffice(name);
    applyOfficeInfo({ office: name, province: '' });
    const now = new Date();
    state.emp = { headers: r.t.emp.headers, rows: r.t.emp.rows, at: now };
    state.orders = { headers: r.t.orders.headers, rows: r.t.orders.rows, wbCol: r.t.orders.wbCol, at: now };
    state.lastRefresh = now;
    log(`⚡ Tải nhanh bằng API: ${fmt(r.t.orders.rows.length)} đơn, ${fmt(r.t.emp.rows.length)} nhân viên (${name}).`);
    state.lastSave = await apiSave(name, r.t, r.items, r.total, r.diag);
    return true;
  }
  async function switchOffice(name) {
    return runBusy(async () => {
      if (state.fast) {
        try { return await apiOverview(name); }
        catch (e) { log(`Tải nhanh bằng API lỗi: ${e.message}. Chuyển sang cách chọn ô Bưu cục trên trang.`, true); }
      }
      return (await selectOffice(name)) ? doUpdate(name) : false;
    });
  }

  // Danh sách bưu cục gợi ý: lấy từ hệ thống (nếu được) + các bưu cục đã dùng gần đây
  const asArray = obj => {
    if (Array.isArray(obj)) return obj;
    if (!obj || typeof obj !== 'object') return [];
    for (const k of ['list', 'records', 'rows', 'items', 'resultList', 'dataList', 'siteList', 'children', 'nodes']) if (Array.isArray(obj[k])) return obj[k];
    return obj.data && typeof obj.data === 'object' ? asArray(obj.data) : [];
  };

  async function apiJson(path, body) {
    const opt = body === undefined
      ? { method: 'GET', credentials: 'include', headers: { Accept: 'application/json' } }
      : { method: 'POST', credentials: 'include', headers: { Accept: 'application/json', 'Content-Type': 'application/json' }, body: JSON.stringify(body) };
    const r = await fetch(location.origin + path, opt);
    if (!r.ok) throw new Error(`HTTP ${r.status}`);
    const j = JSON.parse(await r.text());
    if (j && j.success === false) throw new Error(j.message || j.msg || 'API báo thất bại');
    return j;
  }

  function renderOfficeList() {
    const sel = $('#b8s-office-in');
    if (!sel) return;
    const cur = OFFICES.find(n => plain(n) === plain(state.office));
    const extra = state.office && !cur ? [state.office] : [];
    sel.innerHTML = `<option value="">${OFFICES.length || extra.length ? '— Chọn bưu cục —' : '— Chưa có bưu cục: bấm “Chuyển bưu cục” để tự tìm —'}</option>` +
      [...OFFICES, ...extra].map(n => `<option value="${esc(n)}">${esc(n)}</option>`).join('');
    sel.value = cur || (extra[0] || '');
  }

  /* ------------------------------------------------------------------ *
   * Lịch sử theo ngày (IndexedDB của trình duyệt, theo từng bưu cục)
   * ------------------------------------------------------------------ */
  // <idb>
  function makeStore(factory) {
    let dbp = null;
    const open = () => dbp || (dbp = new Promise((res, rej) => {
      const rq = factory.open('b8s-history', 1);
      rq.onupgradeneeded = () => { const db = rq.result; if (!db.objectStoreNames.contains('days')) db.createObjectStore('days', { keyPath: 'id' }); };
      rq.onsuccess = () => res(rq.result);
      rq.onerror = () => { dbp = null; rej(rq.error || new Error('Không mở được kho lưu trữ')); };
      rq.onblocked = () => { dbp = null; rej(new Error('Kho lưu trữ đang bị tab khác giữ')); };
    }));
    const tx = async (mode, fn) => {
      const db = await open();
      return new Promise((res, rej) => {
        const t = db.transaction('days', mode);
        let out; const r = fn(t.objectStore('days'));
        if (r) r.onsuccess = () => { out = r.result; };
        t.oncomplete = () => res(out);
        t.onerror = t.onabort = () => rej(t.error || new Error('Lỗi ghi kho lưu trữ'));
      });
    };
    return {
      all: () => tx('readonly', s => s.getAll()),
      put: rec => tx('readwrite', s => s.put(rec)),
      putMany: recs => tx('readwrite', s => { recs.forEach(r => s.put(r)); }),
      del: id => tx('readwrite', s => s.delete(id))
    };
  }
  // Kho riêng cho bản chụp theo mốc giờ (cơ sở dữ liệu khác, không ảnh hưởng kho lịch sử theo ngày)
  function makeSnapStore(factory) {
    let dbp = null;
    const open = () => dbp || (dbp = new Promise((res, rej) => {
      const rq = factory.open('b8s-snaps', 1);
      rq.onupgradeneeded = () => { const db = rq.result; if (!db.objectStoreNames.contains('snaps')) db.createObjectStore('snaps', { keyPath: 'id' }); };
      rq.onsuccess = () => res(rq.result);
      rq.onerror = () => { dbp = null; rej(rq.error || new Error('Không mở được kho bản chụp')); };
      rq.onblocked = () => { dbp = null; rej(new Error('Kho bản chụp đang bị tab khác giữ')); };
    }));
    const tx = async (mode, fn) => {
      const db = await open();
      return new Promise((res, rej) => {
        const t = db.transaction('snaps', mode);
        let out; const r = fn(t.objectStore('snaps'));
        if (r) r.onsuccess = () => { out = r.result; };
        t.oncomplete = () => res(out);
        t.onerror = t.onabort = () => rej(t.error || new Error('Lỗi ghi kho bản chụp'));
      });
    };
    return {
      all: () => tx('readonly', st => st.getAll()),
      putMany: recs => tx('readwrite', st => { recs.forEach(r => st.put(r)); }),
      delMany: ids => tx('readwrite', st => { ids.forEach(i => st.delete(i)); })
    };
  }
  // </idb>
  const store = (() => { try { return makeStore(indexedDB); } catch { return null; } })();

  async function loadHistory() {
    if (!store) { state.histErr = 'Trình duyệt không hỗ trợ IndexedDB nên không lưu được lịch sử.'; renderUI(); return; }
    try {
      const list = await store.all();
      state.hist = new Map();
      for (const r of list || []) { const n = normalizeRecord(r); if (n) state.hist.set(n.id, n); }
      state.histReady = true; state.histErr = '';
    } catch (e) { state.histErr = `Không đọc được lịch sử: ${e.message}`; }
    try { if (navigator.storage && navigator.storage.persist) navigator.storage.persist(); } catch { /* bỏ qua */ }
    renderUI();
  }

  async function putRecords(recs) {
    if (!store) throw new Error('Không có kho lưu trữ');
    await store.putMany(recs);
    recs.forEach(r => state.hist.set(r.id, r));
  }

  /* ------------------------------------------------------------------ *
   * Bản chụp theo mốc giờ cố định (tab ANM)
   * Đến mốc 10:30, 13:30, 15:30, 18:30, 23:30 (giờ máy) mà trang 800BEST đang mở: tự tải số liệu các bưu cục ANM bằng API
   * và lưu lại. Tải tay hay tự cập nhật khi vào tab ANM chỉ để xem, không tạo bản chụp.
   * ------------------------------------------------------------------ */
  const snapStore = (() => { try { return makeSnapStore(indexedDB); } catch { return null; } })();
  const TAB_ID = Math.random().toString(36).slice(2);
  const API_DET = '/v9-order/amTodayDispBoard/getTodayDetail', API_EMP = '/v9-order/amTodayDispBoard/getTodayEmployeeSum', API_SITE = '/sea-man/node/pageQuery';
  async function apiPost(path, body) {
    // Trang đăng nhập bằng cookie nên header token thường RỖNG; dùng bộ tiêu đề mặc định rồi ghi đè bằng giá trị đã bắt được từ trang
    const hdr = Object.assign({ Accept: 'application/json, text/plain, */*', 'lang-type': 'vi-VN', langType: 'vi-VN', 'TIME-ZONE': 'GMT+7', 'data-masking-N': 'mark', 'no-offset-timeZone': 'true', token: '', 'X-Requested-With': 'XMLHttpRequest' }, state.apiHdr);
    for (const k of Object.keys(hdr)) if (typeof hdr[k] === 'boolean') hdr[k] = String(hdr[k]);
    hdr['Content-Type'] = 'application/json;charset=UTF-8';
    const r = await fetch(location.origin + path, { method: 'POST', credentials: 'include', headers: hdr, body: JSON.stringify(body) });
    if (!r.ok) throw new Error(`HTTP ${r.status}`);
    const j = JSON.parse(await r.text());
    if (!j || j.success === false) throw new Error((j && (j.errMsg || j.message || j.msg)) || 'API báo thất bại');
    return j.data;
  }
  /* ------------------------------------------------------------------ *
   * Dự báo POD2400 cuối ngày bằng AI (tùy chọn, bật/tắt dưới bảng xếp hạng)
   * Chỉ gửi các con số tổng hợp của bưu cục (không có mã đơn, tên ship, thông tin người nhận).
   * ------------------------------------------------------------------ */
  function loadSites() { try { const o = JSON.parse(localStorage.getItem('b8s_sites') || '{}'); if (o && typeof o === 'object') Object.assign(state.sites, o); } catch { /* bỏ qua */ } }
  function saveSites() { try { localStorage.setItem('b8s_sites', JSON.stringify(state.sites)); } catch { /* bỏ qua */ } }
  // ----- Danh sách bưu cục quản lý: thêm / xóa ở Cài đặt, lưu trên trình duyệt -----
  const offHidden = new Set();   // bưu cục đã xóa: ẩn cả khỏi lịch sử cũ
  function loadOffices() {
    try { const a = JSON.parse(localStorage.getItem('b8s_offices') || 'null'); if (Array.isArray(a) && a.every(x => typeof x === 'string' && x.trim())) OFFICES.splice(0, OFFICES.length, ...a.map(x => x.trim())); } catch { /* bỏ qua */ }
    try { const h = JSON.parse(localStorage.getItem('b8s_offhide') || '[]'); if (Array.isArray(h)) h.forEach(x => typeof x === 'string' && offHidden.add(x)); } catch { /* bỏ qua */ }
  }
  function saveOffices() {
    try { localStorage.setItem('b8s_offices', JSON.stringify(OFFICES)); localStorage.setItem('b8s_offhide', JSON.stringify([...offHidden])); } catch { /* bỏ qua */ }
  }
  const siteNm = x => { const v = x.nameLocale; if (typeof v === 'string') return v; if (v && typeof v === 'object') return Object.values(v).find(t => typeof t === 'string') || ''; return x.name || ''; };
  // Kiểm tra tên bưu cục phải ĐÚNG TUYỆT ĐỐI (kể cả chữ hoa, chữ thường, dấu). Trả về {ok, name, id, code} hoặc {ok:false, msg}
  async function officeLookup(typed) {
    const name = String(typed ?? '').trim().replace(/\s+/g, ' ');
    if (!name) return { ok: false, msg: 'Hãy gõ tên bưu cục.' };
    if (OFFICES.includes(name)) return { ok: false, msg: `“${name}” đã có trong danh sách.` };
    let seen = null;
    for (const auth of [true, false]) {
      let list;
      try { list = asArray(await apiPost(API_SITE, { pageNumber: 1, pageSize: 200, fuzzyNameOrCode: name, limitSiteByAuth: auth, statusList: ['ENABLE', 'SHUTDOWN'] })); }
      catch (e) { return { ok: false, msg: `Không kiểm tra được tên bưu cục (${e && e.message || e}). Thử lại sau.` }; }
      const hits = list.filter(x => siteNm(x).trim() === name);
      const hit = hits.find(x => x.status !== 'SHUTDOWN') || hits[0];
      if (hit) { seen = { hit, auth }; break; }
    }
    if (!seen) return { ok: false, msg: `Bưu cục “${name}” không tồn tại. Gõ đúng từng chữ, kể cả chữ hoa, chữ thường và dấu.` };
    const x = seen.hit;
    if (x.status === 'SHUTDOWN') return { ok: false, msg: `Bưu cục “${name}” đã đóng, không thể thêm.` };
    if (!x.id || !x.code) return { ok: false, msg: `Bưu cục “${name}” thiếu mã trong hệ thống, không thể thêm.` };
    if (!seen.auth) return { ok: false, msg: `Bưu cục “${name}” có trên hệ thống nhưng tài khoản của bạn không có quyền xem.` };
    return { ok: true, name, id: x.id, code: String(x.code) };
  }
  const offUi = { busy: false, msg: '', bad: false };
  function offSay(msg, bad) { offUi.msg = msg; offUi.bad = !!bad; const el = $('#b8s-of-st'); if (el) { el.textContent = msg; el.className = bad ? 'bad' : 'muted'; el.style.fontWeight = bad ? '700' : ''; } }
  function renderOffices() {
    const box = $('#b8s-of-list'); if (!box) return;
    const ids = codIds();
    const sig = JSON.stringify([OFFICES, OFFICES.map(o => ids[plain(o)] ?? null), offUi.busy]);
    if (box.dataset.sig !== sig) {
      box.dataset.sig = sig;
      box.innerHTML = OFFICES.map(n => { const v = ids[plain(n)], miss = !Number.isFinite(v);
        return `<div class="row" style="gap:8px;align-items:center;padding:3px 0;border-top:1px solid #eef2f7"><b style="min-width:170px">${esc(n)}</b><label class="muted">Mã COD <input type="number" data-cod="${esc(n)}" value="${miss ? '' : v}" placeholder="chưa có" style="width:120px;${miss ? 'border:1px solid #dc2626' : ''}"></label>${miss ? '<span class="bad">thiếu mã COD: bảng COD của bưu cục này sẽ trống</span>' : ''}<button class="b8sbtn" data-del="${esc(n)}" style="margin-left:auto" ${offUi.busy ? 'disabled' : ''}>Xóa</button></div>`; }).join('');
    }
    const add = $('#b8s-of-add'), name = $('#b8s-of-name'); if (add) add.disabled = offUi.busy; if (name) name.disabled = offUi.busy;
    const st = $('#b8s-of-st'); if (st && st.textContent !== offUi.msg) offSay(offUi.msg, offUi.bad);
  }
  async function officesAdd() {
    if (offUi.busy) return;
    const inp = $('#b8s-of-name'); if (!inp) return;
    offUi.busy = true; offSay('Đang kiểm tra tên bưu cục…'); renderOffices();
    try {
      const r = await officeLookup(inp.value);
      if (!r.ok) { offSay(r.msg, true); return; }
      // Chạy thử lấy đơn: bưu cục không có đơn nào hôm nay thì coi là bưu cục rỗng, không thêm
      offSay(`Tên đúng. Đang chạy thử lấy dữ liệu của “${r.name}”…`);
      state.sites[plain(r.name)] = { id: r.id, code: r.code };
      let probe;
      try { probe = await apiFetchOffice(r.name); }
      catch (e) { delete state.sites[plain(r.name)]; offSay(`Không lấy được dữ liệu thử của “${r.name}” (${e && e.message || e}). Chưa thêm.`, true); return; }
      const nOrders = Math.max(probe.items.length, Number(probe.total) || 0), nEmp = probe.t && probe.t.emp && probe.t.emp.rows ? probe.t.emp.rows.length : 0;
      if (nOrders === 0) { delete state.sites[plain(r.name)]; offSay(`Bưu cục “${r.name}” rỗng: ${nOrders} đơn, ${nEmp} nhân viên. Không thêm.`, true); return; }
      saveSites();
      OFFICES.push(r.name); offHidden.delete(r.name); offHidden.delete(plain(r.name)); saveOffices();
      inp.value = '';
      const ids = codIds();
      offSay(`Đã thêm “${r.name}” (${nOrders} đơn, ${nEmp} nhân viên).${Number.isFinite(ids[plain(r.name)]) ? '' : ' Nhập “Mã COD” ở dòng của bưu cục để có bảng COD.'} Bấm “Tải số liệu” ở tab ANM để lấy số.`);
      log(`Đã thêm bưu cục ${r.name}.`);
      if (state.anm.sel) { state.anm.sel.add(r.name); savePrefs(); }
      renderOfficeList();   // ô chọn bưu cục ở tab Tổng quan
      if (state.view === 'anm') renderAnm();
    } finally { offUi.busy = false; const b = $('#b8s-of-list'); if (b) delete b.dataset.sig; renderOffices(); }
  }
  function officesRemove(name) {
    const i = OFFICES.indexOf(name);
    if (i < 0) return;
    if (!window.confirm(`Xóa “${name}” khỏi danh sách theo dõi?\nLịch sử đã lưu không bị xóa, chỉ không còn hiện.`)) return;
    OFFICES.splice(i, 1); offHidden.add(name); saveOffices();
    if (state.anm.sel) { state.anm.sel.delete(name); savePrefs(); }
    offSay(`Đã xóa “${name}”.`); const b = $('#b8s-of-list'); if (b) delete b.dataset.sig; renderOffices(); renderOfficeList();
    if (state.view === 'anm') renderAnm();
  }
  function officesSaveCod(name, val) {
    const v = parseInt(val, 10), k = plain(name);
    if (!(v > 0)) {   // xóa mã đã nhập tay (mã có sẵn trong script vẫn dùng)
      try { const o = JSON.parse(localStorage.getItem('b8s_codsites') || '{}') || {}; delete o[k]; localStorage.setItem('b8s_codsites', JSON.stringify(o)); } catch { /* bỏ qua */ }
    } else saveCodIds({ [k]: v });
    const b = $('#b8s-of-list'); if (b) delete b.dataset.sig; renderOffices();
  }

  // ----- Tự dò bưu cục: chạy khi bấm “Tải số liệu” (ANM) hoặc “Chuyển bưu cục” (Tổng quan) -----
  const DISC_MAX_NEW = 12, DISC_MAX_PROBE = 80;
  const disc = { on: true, busy: false, undo: null, pick: null, note: '', warn: false, empty: {} };   // empty: bưu cục 0 đơn đã kiểm tra trong ngày (không dò lại)
  function loadDisc() {
    try { const o = JSON.parse(localStorage.getItem('b8s_disc') || '{}') || {}; if (typeof o.on === 'boolean') disc.on = o.on; } catch { /* bỏ qua */ }
    try { const n = Number(localStorage.getItem('b8s_owner')); if (Number.isFinite(n) && n > 0) state.ownerSiteId = n; } catch { /* bỏ qua */ }
  }
  function saveDisc() { try { localStorage.setItem('b8s_disc', JSON.stringify({ on: disc.on })); } catch { /* bỏ qua */ } }
  // Lấy toàn bộ danh sách bưu cục tài khoản này được xem. Thử kèm ownerSiteId (nếu đã biết) rồi thử không kèm.
  async function discList() {
    const owners = []; if (Number.isFinite(state.ownerSiteId)) owners.push(state.ownerSiteId); owners.push(null);
    let lastErr = null;
    for (const ow of owners) {
      try {
        const all = []; let total = Infinity;
        for (let pg = 1; pg <= 20 && all.length < total; pg++) {
          const body = { pageNumber: pg, pageSize: 200, fuzzyNameOrCode: '', limitSiteByAuth: true, statusList: ['ENABLE', 'SHUTDOWN'] };
          if (ow !== null) body.ownerSiteId = ow;
          const d = await apiPost(API_SITE, body), arr = asArray(d);
          const t = Number(d && d.total); total = Number.isFinite(t) && t >= 0 ? t : arr.length;
          if (!arr.length) break;
          all.push(...arr);
        }
        if (all.length) return all;
        lastErr = new Error('hệ thống trả danh sách trống');
      } catch (e) { lastErr = e; }
    }
    throw lastErr || new Error('không lấy được danh sách');
  }
  // Số đơn hôm nay của một bưu cục (gọi nhỏ, chỉ lấy tổng)
  async function discProbe(c) {
    const d = await apiPost(API_DET, { isContainsSecSite: true, siteId: c.id, siteCode: c.code, shouldDispType: '1', searchType: '1', pageSize: 20, currentPage: 1 });
    return Math.max(asArray(d).length, Number(d && d.total) || 0);
  }
  function discRender() {
    const box = $('#b8s-disc'); if (!box) return;
    let h = '';
    if (disc.pick) {
      h = `<div class="al a"><b>Tìm thấy ${disc.pick.length} bưu cục có đơn.</b> Nhiều hơn ${DISC_MAX_NEW} nên tool không tự thêm. Tích những bưu cục bạn muốn theo dõi:<div style="max-height:150px;overflow:auto;margin:4px 0">${disc.pick.map((c, i) => `<label style="display:inline-block;min-width:230px"><input type="checkbox" data-dpick="${i}" ${c.on === false ? '' : 'checked'}> ${esc(c.name)} <span class="muted">(${c.n} đơn)</span></label>`).join('')}</div><button class="b8sbtn solid" data-d="padd">Thêm các bưu cục đã tích</button> <button class="b8sbtn" data-d="pall">Tích hết</button> <button class="b8sbtn" data-d="pnone">Bỏ tích hết</button> <button class="b8sbtn" data-d="pskip">Bỏ qua</button></div>`;
    } else if (disc.note) {
      h = `<div class="al ${disc.warn ? 'a' : 'ok'}" style="display:flex;gap:8px;align-items:center"><span style="flex:1">${esc(disc.note)}</span>${disc.undo && disc.undo.length ? '<button class="b8sbtn" data-d="undo" title="Bỏ các bưu cục vừa được tự thêm">Hoàn tác</button>' : ''}<button class="b8sbtn ghost" data-d="close">×</button></div>`;
    }
    box.style.display = h ? 'block' : 'none';
    if (box.dataset.h !== h) { box.dataset.h = h; box.innerHTML = h; }
  }
  function discAdd(cands) {
    const names = [];
    for (const c of cands) {
      if (OFFICES.includes(c.name)) continue;
      state.sites[plain(c.name)] = { id: c.id, code: c.code };
      OFFICES.push(c.name); offHidden.delete(c.name); offHidden.delete(plain(c.name)); names.push(c.name);
      if (state.anm.sel) state.anm.sel.add(c.name);
    }
    if (names.length) { saveSites(); saveOffices(); savePrefs(); }
    if (names.length) { const b = $('#b8s-of-list'); if (b) delete b.dataset.sig; renderOffices(); renderOfficeList(); }
    return names;
  }
  function discUndo() {
    const names = (disc.undo || []).slice();
    for (const n of names) { const i = OFFICES.indexOf(n); if (i >= 0) OFFICES.splice(i, 1); offHidden.add(n); if (state.anm.sel) state.anm.sel.delete(n); }
    if (names.length) { saveOffices(); savePrefs(); }
    disc.undo = null; disc.note = names.length ? `Đã hoàn tác: bỏ ${names.join(', ')} khỏi danh sách theo dõi.` : ''; disc.warn = false;
    log(disc.note);
    const b = $('#b8s-of-list'); if (b) delete b.dataset.sig;
    renderOffices(); renderOfficeList(); discRender(); if (state.view === 'anm') renderAnm();
  }
  function discPickDone(addIt) {
    const p = disc.pick || []; disc.pick = null;
    const chosen = addIt ? p.filter(c => c.on !== false) : [];
    const names = discAdd(chosen);
    disc.undo = names.length ? names : null;
    disc.note = names.length ? `Đã thêm ${names.length} bưu cục: ${names.join(', ')}. Bấm “Tải số liệu” lại để lấy số.` : 'Không thêm bưu cục nào.'; disc.warn = !names.length;
    log(disc.note); discRender(); if (state.view === 'anm') renderAnm();
  }
  // Trả về {added, ...} hoặc null nếu không chạy. Không bao giờ ném lỗi: lỗi chỉ được báo, danh sách cũ giữ nguyên.
  async function discoverOffices() {
    if (!disc.on || disc.busy) return null;
    disc.busy = true; disc.pick = null;
    const st = (m) => { log(m); };
    try {
      st('Đang dò các bưu cục bạn quản lý…');
      let list;
      try { list = await discList(); }
      catch (e) {
        const m = `Không dò được danh sách bưu cục (${e && e.message || e}). Giữ nguyên danh sách hiện tại${OFFICES.length ? '' : '; bạn có thể thêm bưu cục thủ công ở Cài đặt'}.`;
        disc.note = m; disc.warn = true; disc.undo = null; log(m, true); discRender(); return null;
      }
      const pk = discPick(list, new Set(OFFICES.map(plainOf)), offHidden);
      const today = dateKey(new Date());
      const todo = pk.cand.filter(c => disc.empty[c.id] !== today);
      const left = Math.max(0, todo.length - DISC_MAX_PROBE), batch = todo.slice(0, DISC_MAX_PROBE);
      const ok = [], empty = pk.cand.filter(c => disc.empty[c.id] === today).map(c => c.name), failed = [];
      let p = 0;
      const worker = async () => {
        while (p < batch.length) {
          const c = batch[p++];
          try { const n = await discProbe(c); if (n > 0) ok.push(Object.assign({}, c, { n })); else { empty.push(c.name); disc.empty[c.id] = today; } }
          catch { failed.push(c.name); }
        }
      };
      await Promise.all([worker(), worker(), worker()]);
      ok.sort((a, b) => a.name.localeCompare(b.name, 'vi'));
      let added = [];
      if (ok.length > DISC_MAX_NEW) { disc.pick = ok; disc.note = ''; disc.undo = null; }
      else { added = discAdd(ok); disc.undo = added.length ? added : null; }
      const quiet = !added.length && !empty.length && !failed.length && !left && !disc.pick && OFFICES.length;
      const sum = discSummary(pk, added, empty, failed, left);
      if (!disc.pick) {
        disc.note = quiet ? '' : (!OFFICES.length ? `${sum} Chưa có bưu cục nào có đơn hôm nay.` : sum);
        disc.warn = !added.length;
      }
      if (!quiet) log(sum, !!failed.length);
      discRender();
      if (added.length && state.view === 'anm') renderAnm();
      return { added, empty, failed, picked: !!disc.pick };
    } catch (e) {
      log(`Dò bưu cục lỗi: ${e && e.message || e}`, true); return null;
    } finally { disc.busy = false; }
  }
  async function resolveSite(name) {
    const key = plain(name);
    if (state.sites[key]) return state.sites[key];
    const nm = x => { const v = x.nameLocale; if (typeof v === 'string') return v; if (v && typeof v === 'object') return Object.values(v).find(t => typeof t === 'string') || ''; return x.name || ''; };
    let hit = null, n = 0;
    for (const auth of [true, false]) {
      const list = asArray(await apiPost(API_SITE, { pageNumber: 1, pageSize: 200, fuzzyNameOrCode: name, limitSiteByAuth: auth, statusList: ['ENABLE', 'SHUTDOWN'] }));
      n = list.length;
      hit = list.find(x => plain(nm(x)) === key && x.status !== 'SHUTDOWN') || list.find(x => plain(nm(x)) === key);
      if (hit) break;
    }
    if (!hit || !hit.id || !hit.code) throw new Error(`không tìm thấy mã bưu cục “${name}” (API trả ${n} kết quả)`);
    state.sites[key] = { id: hit.id, code: String(hit.code) };
    saveSites();
    return state.sites[key];
  }
  // Gọi API lấy toàn bộ đơn + nhân viên của một bưu cục và dựng thành bảng giống trên trang
  async function apiFetchOffice(name) {
    const site = await resolveSite(name);
    const base = { isContainsSecSite: true, siteId: site.id, siteCode: site.code, shouldDispType: '1' };
    const det = Object.assign({}, base, { searchType: '1' });
    const sizes = state.apiPs ? [state.apiPs] : [1000, 500, 200, 100, 50];
    const empP = apiPost(API_EMP, base);
    empP.catch(() => { /* lỗi sẽ được báo khi đợi kết quả bên dưới; tránh báo lỗi chưa xử lý nếu API đơn lỗi trước */ });
    let first = null, lastErr = null;
    for (const ps of sizes) {
      try { first = await apiPost(API_DET, Object.assign({}, det, { pageSize: ps, currentPage: 1 })); state.apiPs = ps; break; }
      catch (e) { lastErr = e; if (state.apiPs) state.apiPs = 0; }
    }
    if (!first) { if (state.apiPs === 0) { state.apiPs = null; } throw lastErr || new Error('API đơn không trả dữ liệu'); }
    const empData = await empP;
    let items = ((first && first.list) || []).slice();
    const total = Number(first && first.total) || 0, per = items.length;
    let calls = 1, dropped = 0;
    if (per && items.length < total) {
      const pages = Math.min(400, Math.ceil(total / per));
      const nums = []; for (let p = 2; p <= pages; p++) nums.push(p);
      for (let i = 0; i < nums.length; i += 6) {
        const got = await Promise.all(nums.slice(i, i + 6).map(p => apiPost(API_DET, Object.assign({}, det, { pageSize: per, currentPage: p })).then(d => (d && d.list) || [])));
        got.forEach(g => items.push(...g)); calls += got.length;
      }
      // Nhiều trang: dữ liệu đổi giữa các lần gọi có thể làm một dòng xuất hiện ở hai trang; chỉ loại các dòng GIỐNG HỆT nhau hoàn toàn
      const sg = new Set(), uniq = [];
      for (const it of items) { let k; try { k = JSON.stringify(it); } catch { k = null; } if (k !== null) { if (sg.has(k)) { dropped++; continue; } sg.add(k); } uniq.push(it); }
      items = uniq;
    }
    const t = apiToTables(items, asArray(empData));
    const diag = { rows: items.length, pages: calls, dup: t.dupN, empty: t.emptyN, dropped };
    state.sign = { keys: t.signKeys, okN: t.okN, signN: t.signN };
    if (t.okN > 0 && t.signN === 0 && !state.signWarned) { state.signWarned = true; log(`${name}: không thấy trường giờ ký nhận trong đơn nên chưa có “GTC đầu tiên”. Bấm “Chẩn đoán API” rồi gửi dòng signTimeKeys.`, true); }
    if (t.unknown.length) log(`${name}: có mã trạng thái chưa biết (${t.unknown.join(', ')}), tạm xếp vào “khác”.`, true);
    return { t, items, total, diag };
  }
  async function apiLoadOffice(name) {
    const { t, items, total, diag } = await apiFetchOffice(name);
    return apiSave(name, t, items, total, diag);
  }
  async function apiSave(name, t, items, total, diag) {
    const calc = computeDashboard({ orders: t.orders, emp: t.emp, tiktok: parseList(state.tiktok), statusMap: state.statusMap, now: new Date(), rules: state.dm });
    const d = decideSave({ hasOrders: true, pageOffice: name, startOffice: name, expectedOffice: name, ovNeed: total, computedTotal: calc.cards.all.total, diag });
    const rec = buildRecord({ office: canonOffice(name, OFFICES) || name, now: new Date(), calc, verified: !!d.verified });
    if (!d.save) { return { saved: false, reason: d.reason, rec: d.force ? rec : null }; }
    await putRecords([rec]);
    state.lastSaved = new Date();
    log(`Đã lưu lịch sử ${rec.office} ngày ${fmtDMY(rec.date)} (API, ${fmt(items.length)} đơn).`);
    return { saved: true, verified: !!d.verified };
  }

  async function trySaveHistory(expect, startOffice) {
    state.pending = null;
    const c = getCalc();
    // Nếu không đọc được tên bưu cục từ trang thì dùng đúng bưu cục tool vừa chọn (tool đã chọn và chờ trang tải xong)
    const pageOffice = readOfficeInfo().office || expect || state.office || '';
    const m = (document.body?.innerText || '').match(/Tổng đơn cần phát hàng\s*([\d,.]+)/i);
    const d = decideSave({
      hasOrders: !!(c && state.orders), pageOffice, startOffice, expectedOffice: expect,
      ovNeed: m ? parseNum(m[1]) : (state.ov && Number.isFinite(state.ov.need) ? state.ov.need : NaN), computedTotal: c ? c.cards.all.total : 0
    });
    if (!c) return { saved: false, reason: 'Chưa tính được số liệu từ bảng đơn.' };
    const rec = d.reason && !d.force ? null : buildRecord({ office: canonOffice(pageOffice, OFFICES), now: new Date(), calc: c, verified: !!d.verified });
    if (d.save && rec) {
      try {
        await putRecords([rec]);
        state.lastSaved = new Date();
        log(`Đã lưu lịch sử ${rec.office} ngày ${fmtDMY(rec.date)}${d.verified ? '' : ' (chưa đối chiếu được tổng đơn với trang)'}.`);
        return { saved: true, verified: !!d.verified };
      } catch (e) { log(`Không lưu được lịch sử: ${e.message}`, true); return { saved: false, reason: `Lỗi ghi kho lưu trữ: ${e.message}` }; }
    }
    log(`Chưa lưu lịch sử: ${d.reason}`, true);
    if (d.force && rec) state.pending = { reason: d.reason, rec };
    return { saved: false, reason: d.reason, rec: d.force ? rec : null };
  }

  async function forceSave() {
    const p = state.pending;
    if (!p) return;
    try {
      await putRecords([{ ...p.rec, verified: false }]);
      state.lastSaved = new Date(); state.pending = null;
      log(`Đã lưu lịch sử ${p.rec.office} ngày ${fmtDMY(p.rec.date)} (lưu thủ công).`);
    } catch (e) { log(`Không lưu được lịch sử: ${e.message}`, true); }
    renderUI();
  }

  function backupInfo() {
    let last = 0; try { last = Number(localStorage.getItem('b8s_lastBackup') || 0); } catch { /* bỏ qua */ }
    return { last, due: backupDue(state.hist.size, last, Date.now()) };
  }

  function downloadText(name, text, type) {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([text], { type }));
    a.download = name; document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 4000);
  }
  function backupHistory() {
    const records = [...state.hist.values()];
    if (!records.length) { log('Chưa có lịch sử để sao lưu.', true); return; }
    downloadText(`800BEST_LichSu_${dateKey(new Date())}.json`, JSON.stringify({ app: '800best-history', version: 1, exportedAt: new Date().toISOString(), records, snaps: state.snap.list }), 'application/json');
    try { localStorage.setItem('b8s_lastBackup', String(Date.now())); } catch { /* bỏ qua */ }
    log(`Đã tạo file sao lưu ${records.length} bản ghi.`);
    renderUI();
  }
  async function restoreHistory(file) {
    let obj;
    try { obj = JSON.parse(await file.text()); } catch { log('File không đọc được (không phải JSON).', true); return; }
    const v = validateBackup(obj);
    if (!v.ok) { log(v.error, true); return; }
    // bản chụp theo mốc giờ (nếu file có): chỉ thêm những bản chưa có hoặc cũ hơn
    const snapsIn = (Array.isArray(obj.snaps) ? obj.snaps : []).map(normalizeSnap).filter(Boolean);
    const snapPut = snapsIn.filter(r => { const cur = snapAt(state.snap.idx, r.office, r.date, r.slot); return !cur || String(r.runAt) > String(cur.runAt); });
    if (snapPut.length && snapStore) {
      try { await snapStore.putMany(snapPut); await loadSnaps(); log(`Đã khôi phục ${snapPut.length} bản chụp theo mốc giờ.`); }
      catch (e) { log(`Không khôi phục được bản chụp theo mốc giờ: ${e.message}`, true); }
    }
    const plan = planMerge(state.hist, v.records);
    if (!plan.put.length) { log(`Không có gì cần khôi phục (${plan.kept} bản ghi đã có bản mới hơn hoặc bằng).`); return; }
    if (!confirm(`Khôi phục lịch sử?\nThêm mới: ${plan.added} ngày\nCập nhật (bản trong file mới hơn): ${plan.updated} ngày\nGiữ nguyên (đã có bản mới hơn): ${plan.kept} ngày`)) return;
    try { await putRecords(plan.put); try { localStorage.setItem('b8s_lastBackup', String(Date.now())); } catch { /* bỏ qua */ } log(`Đã khôi phục: thêm ${plan.added}, cập nhật ${plan.updated}, giữ nguyên ${plan.kept}.`); }
    catch (e) { log(`Khôi phục thất bại: ${e.message}`, true); }
    renderUI();
  }

  /* ---- Tab Lịch sử ---- */
  const HCOLS = [
    { k: 'total', label: 'Tổng đơn', kind: 'avg' },
    { k: 'scanRate', label: 'Tỷ lệ quét giao', kind: 'pct' },
    { k: 'pod', label: 'POD2400', kind: 'pct', n: 'podN' },
    { k: 'podTT', label: 'POD TTS', kind: 'pct', n: 'podTTN' },
    { k: 'gtcTrip', label: 'GTC chuyến', kind: 'pct' },
    { k: 'gtcTotal', label: 'GTC Total', kind: 'pct' },
    { k: 'backlog', label: 'Backlog 3 ngày', kind: 'avg', last: 'backlogLast' },
    { k: 'odr', label: 'TTS đến hạn ODR', kind: 'avg', last: 'odrLast' },
    { k: 'nWork', label: 'Shipper đi làm', kind: 'avg' }
  ];

  const histOffices = () => {
    const set = new Set(OFFICES);
    for (const r of state.hist.values()) { const o = canonOffice(r.office, OFFICES); if (!offHidden.has(o)) set.add(o); }
    return [...set];
  };
  function histFilter() {
    const hv = state.hv, today = dateKey(new Date());
    const office = hv.office || (state.office ? canonOffice(state.office, OFFICES) : '*');
    const rg = hv.preset === 'custom' ? { from: hv.from || '0000-01-01', to: hv.to || '9999-12-31' } : rangeFor(hv.preset, today);
    const recs = [...state.hist.values()]
      .filter(r => (office === '*' || sameOffice(r.office, office)) && r.date >= rg.from && r.date <= rg.to)
      .sort((a, b) => a.date < b.date ? -1 : a.date > b.date ? 1 : a.office < b.office ? -1 : 1);
    return { office, rg, recs };
  }
  function trendRows(recs, group) { return withDeltas(groupRecords(recs, group)); }

  function hcell(col, m, delta, multi) {
    const v = m[col.k];
    if (v === null || v === undefined || !Number.isFinite(v)) return '<td>—</td>';
    if (col.kind === 'pct') {
      const n = col.n && m[col.n], tip = n ? `${fmt(n[0])} / ${fmt(n[1])} đơn` : '';
      const d = delta && Number.isFinite(delta[col.k]) ? delta[col.k] : null;
      const dd = d === null || Math.abs(d) < 0.05 ? '' : `<small class="${d > 0 ? 'up' : 'dn'}">${d > 0 ? '▲' : '▼'}${round1(Math.abs(d)).toLocaleString('vi-VN')}</small>`;
      return `<td class="${rateClass(v)}" title="${esc(tip)}">${esc(fmtVal(v, 'pct'))}${dd}</td>`;
    }
    const last = col.last && multi && m[col.last] !== null && m[col.last] !== undefined ? `Ngày cuối kỳ: ${fmt(m[col.last])}` : '';
    return `<td title="${esc(last)}">${esc(fmtAvg(v))}</td>`;
  }

  function renderHistory() {
    const box = $('#b8s-hist');
    if (!box) return;
    const hv = state.hv, f = histFilter(), today = dateKey(new Date());
    const offs = histOffices();
    const bi = backupInfo();
    const opt = (v, cur, l) => `<option value="${esc(v)}" ${v === cur ? 'selected' : ''}>${esc(l)}</option>`;
    const cmpDate = hv.cmpDate || ([...state.hist.values()].map(r => r.date).sort().pop() || today);
    const controls = `
      <div class="b8s-panel">
        <div class="row">
          <label>Bưu cục<br><select data-h="office">${opt('*', f.office, 'Tất cả (gộp)')}${offs.map(o => opt(o, f.office, o)).join('')}</select></label>
          <label>Xem<br><select data-h="mode">${opt('trend', hv.mode, 'Xu hướng theo kỳ')}${opt('ship', hv.mode, 'Shipper (tổng hợp)')}${opt('cmp', hv.mode, 'So sánh các bưu cục')}</select></label>
          ${hv.mode === 'cmp' ? `<label>Ngày<br><input type="date" data-h="cmpDate" value="${esc(cmpDate)}"></label>` : `
          <label>Khoảng thời gian<br><select data-h="preset">${[['d7', '7 ngày gần nhất'], ['d30', '30 ngày gần nhất'], ['wk', 'Tuần này'], ['lastwk', 'Tuần trước'], ['mo', 'Tháng này'], ['lastmo', 'Tháng trước'], ['all', 'Tất cả'], ['custom', 'Tự chọn…']].map(([v, l]) => opt(v, hv.preset, l)).join('')}</select></label>
          ${hv.preset === 'custom' ? `<label>Từ<br><input type="date" data-h="from" value="${esc(hv.from)}"></label><label>Đến<br><input type="date" data-h="to" value="${esc(hv.to)}"></label>` : ''}
          ${hv.mode === 'trend' ? `<label>Gộp theo<br><select data-h="group">${[['day', 'Ngày'], ['week', 'Tuần'], ['month', 'Tháng']].map(([v, l]) => opt(v, hv.group, l)).join('')}</select></label>` : ''}`}
          <span style="flex:1"></span>
          <button class="b8sbtn solid" data-h-act="backup">Sao lưu</button>
          <button class="b8sbtn solid" data-h-act="restore">Khôi phục</button>
          <input type="file" id="b8s-restore-file" accept=".json,application/json" style="display:none">
        </div>
        ${bi.due ? '<div class="b8s-warn" style="margin:6px 0 0">⚠ Đã hơn 7 ngày chưa sao lưu lịch sử. Bấm “Sao lưu” để tải file về máy (phòng khi xóa dữ liệu trình duyệt).</div>' : ''}
        ${state.histErr ? `<div class="b8s-warn" style="margin:6px 0 0">${esc(state.histErr)}</div>` : ''}
      </div>`;

    let body = '';
    if (!state.hist.size) {
      body = `<div class="b8s-empty"><b>Chưa có lịch sử</b><br>Mỗi lần bạn chọn bưu cục và tải xong số liệu, chỉ số của ngày hôm đó được tự lưu (chỉ số tổng và từng shipper, không lưu mã vận đơn hay thông tin người nhận). Tải lại trong ngày thì bản lưu của ngày đó được cập nhật.</div>`;
    } else if (hv.mode === 'cmp') {
      const recs = [...state.hist.values()].filter(r => r.date === cmpDate).sort((a, b) => a.office < b.office ? -1 : 1);
      if (!recs.length) body = `<div class="b8s-empty">Không có bản ghi nào của ngày ${esc(fmtDMY(cmpDate))}.</div>`;
      else {
        const rows = recs.map(r => ({ r, m: metricsOf(aggregate([r])) }));
        const tot = metricsOf(aggregate(recs));
        body = `<div class="b8s-tw"><table><thead><tr><th class="l g-nav">Bưu cục</th>${HCOLS.map(c => `<th class="g-gh">${esc(c.label)}</th>`).join('')}</tr></thead><tbody>${
          rows.map(({ r, m }) => `<tr><td class="l"><b>${esc(r.office)}</b>${r.verified ? '' : ' <span title="Chưa đối chiếu được tổng đơn với trang">⚠</span>'}</td>${HCOLS.map(c => hcell(c, m, null, false)).join('')}</tr>`).join('')
        }${recs.length > 1 ? `<tr class="total"><td class="l">TỔNG</td>${HCOLS.map(c => hcell(c, tot, null, false)).join('')}</tr>` : ''}</tbody></table></div>`;
      }
    } else if (hv.mode === 'ship') {
      let list = aggregateShippers(f.recs);
      const sk = hv.ssort.key, dir = hv.ssort.dir;
      list.sort((a, b) => {
        const x = a[sk], y = b[sk];
        if (typeof x === 'string' || typeof y === 'string') return String(x ?? '').localeCompare(String(y ?? ''), 'vi') * dir;
        const nx = Number.isFinite(x) ? x : -Infinity, ny = Number.isFinite(y) ? y : -Infinity;
        return (nx - ny) * dir;
      });
      const cols = [['name', 'Shipper', 'text'], ['office', 'Bưu cục', 'text'], ['days', 'Số ngày đi làm', 'int'], ['sc', 'Đã quét giao', 'int'], ['avgSc', 'TB/ngày', 'avg'], ['ok', 'Giao thành công', 'int'], ['rateTrip', '% GTC chuyến', 'pct'], ['tt', 'Đơn TTS', 'int'], ['tk', 'TTS giao thành công', 'int'], ['rateTT', '% GTC TTS', 'pct'], ['sg', 'Đã ký nhận (tab NV)', 'int']];
      body = list.length ? `<div class="b8s-tw"><table><thead><tr><th class="g-nav">#</th>${cols.map(([k, l]) => `<th class="g-gh ${k === 'name' ? 'l' : ''}" data-sk="${k}">${esc(l)}${hv.ssort.key === k ? (dir > 0 ? ' ▲' : ' ▼') : ''}</th>`).join('')}</tr></thead><tbody>${
        list.map((x, i) => `<tr><td>${i + 1}</td>${cols.map(([k, , kind]) => kind === 'text' ? `<td class="l">${esc(x[k])}</td>` : kind === 'pct' ? `<td class="${rateClass(x[k])}">${esc(fmtVal(x[k], 'pct'))}</td>` : kind === 'avg' ? `<td>${x[k] === null ? '—' : esc(fmtAvg(x[k]))}</td>` : `<td>${x[k] === null ? '—' : esc(fmt(x[k]))}</td>`).join('')}</tr>`).join('')
      }</tbody></table></div>` : '<div class="b8s-empty">Không có dữ liệu shipper trong khoảng này.</div>';
    } else {
      if (!f.recs.length) body = '<div class="b8s-empty">Không có bản ghi nào trong khoảng thời gian này.</div>';
      else {
        const list = trendRows(f.recs, hv.group), show = [...list].reverse();
        const whole = list.length > 1 ? metricsOf(aggregate(f.recs)) : null;
        const canDel = hv.group === 'day' && f.office !== '*';
        body = `<div class="b8s-tw"><table><thead><tr><th class="l g-nav">Kỳ</th><th class="g-nav">Số ngày</th>${HCOLS.map(c => `<th class="g-gh">${esc(c.label)}</th>`).join('')}${canDel ? '<th class="g-nav"></th>' : ''}</tr></thead><tbody>${
          (whole ? `<tr class="total"><td class="l">Cả kỳ</td><td>${fmt(whole.days)}</td>${HCOLS.map(c => hcell(c, whole, null, true)).join('')}${canDel ? '<td></td>' : ''}</tr>` : '') +
          show.map(g => `<tr><td class="l"><b>${esc(g.label)}</b>${g.recs.some(r => !r.verified) ? ' <span title="Có bản ghi chưa đối chiếu được tổng đơn với trang">⚠</span>' : ''}</td><td>${fmt(g.m.days)}</td>${HCOLS.map(c => hcell(c, g.m, g.delta, g.m.days > 1)).join('')}${canDel ? `<td><button class="b8sx" title="Xóa bản ghi ngày này" data-del="${esc(g.recs[0].id)}">✕</button></td>` : ''}</tr>`).join('')
        }</tbody></table></div>
        <div class="b8s-notes">Tỷ lệ của tuần/tháng được tính lại từ tổng số đơn của cả kỳ (không lấy trung bình các %). Tổng đơn, Backlog, TTS đến hạn ODR, Shipper đi làm là trung bình mỗi ngày; rê chuột vào ô để xem số đơn hoặc giá trị ngày cuối kỳ. ▲▼ là chênh lệch điểm % so với kỳ trước.</div>`;
      }
    }
    box.innerHTML = controls + body;
  }

  function onHistChange(e) {
    const t = e.target, hv = state.hv;
    if (t.id === 'b8s-restore-file') { const f = t.files && t.files[0]; t.value = ''; if (f) restoreHistory(f); return; }
    const k = t.dataset && t.dataset.h;
    if (!k) return;
    hv[k] = t.value;
    if (k === 'preset' && t.value === 'custom' && !hv.from) { const r = rangeFor('d30', dateKey(new Date())); hv.from = r.from; hv.to = r.to; }
    renderHistory();
  }
  async function onHistClick(e) {
    const b = e.target.closest('[data-h-act]'), th = e.target.closest('th[data-sk]'), del = e.target.closest('[data-del]');
    if (b) {
      if (b.dataset.hAct === 'backup') backupHistory();
      else if (b.dataset.hAct === 'restore') { const fi = $('#b8s-restore-file'); if (fi) fi.click(); }
    } else if (th) {
      const key = th.dataset.sk, ss = state.hv.ssort;
      state.hv.ssort = ss.key === key ? { key, dir: -ss.dir } : { key, dir: (key === 'name' || key === 'office') ? 1 : -1 };
      renderHistory();
    } else if (del) {
      const id = del.dataset.del;
      if (!confirm(`Xóa bản ghi ${id.replace('|', ' – ')}?`)) return;
      try { await store.del(id); state.hist.delete(id); log(`Đã xóa bản ghi ${id}.`); } catch (er) { log(`Không xóa được: ${er.message}`, true); }
      renderUI();
    }
  }

  // Thêm các sheet lịch sử vào file Excel (theo bộ lọc đang chọn ở tab Lịch sử)
  function addHistorySheets(addSheet) {
    if (!state.hist.size) return;
    const f = histFilter();
    if (!f.recs.length) return;
    const pct = v => Number.isFinite(v) ? round1(v) : '';
    const avg = v => (v === null || v === undefined || !Number.isFinite(v)) ? '' : round1(v);
    const head = ['Kỳ', 'Số ngày', 'Tổng đơn (TB/ngày)', 'Tỷ lệ quét giao (%)', 'POD2400 (%)', 'POD TTS (%)', 'GTC chuyến (%)', 'GTC Total (%)', 'Backlog 3 ngày (TB/ngày)', 'Backlog ngày cuối', 'TTS đến hạn ODR (TB/ngày)', 'ODR ngày cuối', 'Shipper đi làm (TB/ngày)'];
    const line = (label, m) => [label, m.days, avg(m.total), pct(m.scanRate), pct(m.pod), pct(m.podTT), pct(m.gtcTrip), pct(m.gtcTotal), avg(m.backlog), avg(m.backlogLast), avg(m.odr), avg(m.odrLast), avg(m.nWork)];
    const list = trendRows(f.recs, state.hv.group);
    addSheet('Lich su', [head, ...list.map(g => line(g.label, g.m)), ...(list.length > 1 ? [line('CẢ KỲ', metricsOf(aggregate(f.recs)))] : [])]);
    addSheet('Theo ngay', [['Ngày', 'Bưu cục', 'Tổng đơn', 'Chưa quét giao', 'Giao thành công', 'Đơn mới', 'Đơn mới giao thành công', 'TTS tổng', 'TTS giao thành công', 'Backlog 3 ngày', 'Chờ duyệt', 'Giao thất bại', 'Nhập kho chưa quét', 'TTS đến hạn ODR', 'Shipper đi làm', 'Đã đối chiếu'],
      ...f.recs.map(r => [r.date, r.office, r.all.total, r.all.unscanned, r.all.ok, r.all.newTotal, r.all.newOk, r.tt.total, r.tt.ok, r.backlog ? r.backlog.total : '', r.backlog ? r.backlog.wait : '', r.backlog ? r.backlog.fail : '', r.backlog ? r.backlog.unscanned : '', r.odrTT ?? '', r.nWork, r.verified ? 'Có' : 'Chưa'])]);
    const sh = aggregateShippers(f.recs);
    if (sh.length) {
      addSheet('Shipper tong', [['Shipper', 'Bưu cục', 'Số ngày đi làm', 'Đã quét giao', 'TB/ngày', 'Giao thành công', '% GTC chuyến', 'Đơn TTS', 'TTS giao thành công', '% GTC TTS', 'Đã ký nhận (tab NV)'],
        ...sh.sort((a, b) => (b.rateTrip ?? -1) - (a.rateTrip ?? -1)).map(x => [x.name, x.office, x.days, x.sc, avg(x.avgSc), x.ok, pct(x.rateTrip), x.tt, x.tk, pct(x.rateTT), x.sg ?? ''])]);
      const day = [['Ngày', 'Bưu cục', 'Shipper', 'Đã quét giao', 'Giao thành công', '% GTC chuyến', 'Đơn TTS', 'TTS giao thành công', '% GTC TTS', 'Đã ký nhận (tab NV)']];
      for (const r of f.recs) for (const s of r.ships) day.push([r.date, r.office, s.n, s.sc, s.ok, pct(ratio(s.ok, s.sc)), s.tt, s.tk, pct(ratio(s.tk, s.tt)), s.sg ?? '']);
      addSheet('Shipper theo ngay', day);
    }
  }

  /* ------------------------------------------------------------------ *
   * Tab ANM: so sánh nhiều bưu cục (đọc từ lịch sử đã lưu)
   * ------------------------------------------------------------------ */
  const COD_IDS_BUILTIN = {};
  function codIds() {
    const m = {}; for (const [k, v] of Object.entries(COD_IDS_BUILTIN)) m[plain(k)] = v;
    try { const o = JSON.parse(localStorage.getItem('b8s_codsites') || '{}'); if (o && typeof o === 'object') for (const [k, v] of Object.entries(o)) if (Number.isFinite(+v)) m[k] = +v; } catch { /* bỏ qua */ }
    return m;
  }
  function saveCodIds(learned) {
    try { const o = JSON.parse(localStorage.getItem('b8s_codsites') || '{}') || {}; localStorage.setItem('b8s_codsites', JSON.stringify(Object.assign(o, learned))); } catch { /* bỏ qua */ }
  }
  const isFS = o => /^\s*fs(\s|$)/i.test(String(o ?? ''));   // bưu cục FS không theo dõi COD
  function renderThresholds() {
    const box = $('#b8s-th');
    if (!box) return;
    const row = (k, lab, fields) => `<tr><td class="l">${esc(lab)}</td>${fields.map(([f, l]) => `<td>${esc(l)} <input type="number" step="any" style="width:70px" data-th="${k}.${f}" value="${Object.assign({}, ANM_TH[k], state.th[k])[f]}"></td>`).join('')}</tr>`;
    box.innerHTML = `<table><tbody>${ANM_METRICS.filter(m => ANM_TH[m.k]).map(m => ANM_TH[m.k].lo !== undefined
      ? row(m.k, m.label + ' (%)', [['lo', 'Đỏ nếu dưới'], ['hi', 'Xanh từ']])
      : row(m.k, m.label, [['ok', 'Xanh nếu ≤'], ['warn', 'Đỏ nếu >']])).join('')}</tbody></table><div class="row" style="margin-top:8px"><button class="b8sbtn solid" id="b8s-th-reset">Đặt lại mặc định</button></div>`;
  }

  /* ------------------------------------------------------------------ *
   * Lưu cấu hình
   * ------------------------------------------------------------------ */
  function loadPrefs() {
    Object.assign(state.wx, WX_DEFAULT);
    state.dm = dmRulesOf(null);
    try {
      const p = JSON.parse(localStorage.getItem('b8s_prefs') || '{}');
      if (p.wx && typeof p.wx === 'object') {
        if (typeof p.wx.on === 'boolean') state.wx.on = p.wx.on;
        if (Number.isFinite(p.wx.lat) && Math.abs(p.wx.lat) <= 90) state.wx.lat = p.wx.lat;
        if (Number.isFinite(p.wx.lon) && Math.abs(p.wx.lon) <= 180) state.wx.lon = p.wx.lon;
      }
      if (typeof p.tiktok === 'string') state.tiktok = p.tiktok;
      if (p.dm && typeof p.dm === 'object') state.dm = dmRulesOf(p.dm);
      if (p.th && typeof p.th === 'object') state.th = p.th;
      if (p.anm && typeof p.anm === 'object') {
        const a = state.anm;
        if (Array.isArray(p.anm.sel)) a.sel = new Set(p.anm.sel.filter(x => typeof x === 'string'));
        if (CMP_MODES.includes(p.anm.cmp)) a.cmp = p.anm.cmp;
        if (RANK_KEYS.includes(p.anm.rankKey)) a.rankKey = p.anm.rankKey;
        if (RANK_KEYS.includes(p.anm.trendKey)) a.trendKey = p.anm.trendKey;
      }
      if (typeof p.fast === 'boolean') state.fast = p.fast;
      if (p.statusMap && typeof p.statusMap === 'object') state.statusMap = p.statusMap;
      if (Array.isArray(p.recent)) state.recent = p.recent.filter(x => typeof x === 'string').slice(0, 20);
    } catch { /* bỏ qua */ }
  }
  function savePrefs() {
    try {
      localStorage.setItem('b8s_prefs', JSON.stringify({ th: state.th, anm: { sel: state.anm.sel ? [...state.anm.sel] : null, cmp: state.anm.cmp, rankKey: state.anm.rankKey, trendKey: state.anm.trendKey }, tiktok: state.tiktok, dm: state.dm, fast: state.fast, statusMap: state.statusMap, recent: state.recent, wx: { on: state.wx.on, lat: state.wx.lat, lon: state.wx.lon } }));
    } catch { /* bỏ qua */ }
  }

  /* ------------------------------------------------------------------ *
   * Giao diện
   * ------------------------------------------------------------------ */
  function setBusy(b) {
    document.querySelectorAll('#b8s-app .b8sbtn[data-act]').forEach(btn => { btn.disabled = b; btn.style.opacity = b ? .6 : 1; });
  }

  function getCalc() {
    if (!state.orders) return null;
    const key = [state.orders.at.getTime(), state.emp ? state.emp.at.getTime() : 0, state.tiktok, JSON.stringify(state.statusMap), JSON.stringify(state.dm), new Date().toDateString()].join('|');
    if (state.calcKey !== key) {
      state.calc = computeDashboard({ orders: state.orders, emp: state.emp, tiktok: parseList(state.tiktok), statusMap: state.statusMap, now: new Date(), rules: state.dm });
      state.calcKey = key;
    }
    return state.calc;
  }

  function rateClass(v, good) {
    if (typeof v !== 'number' || !Number.isFinite(v)) return '';
    if (good === 'low') return v <= 10 ? 'r-ok' : v <= 20 ? 'r-mid' : 'r-bad';
    return v >= 75 ? 'r-ok' : v >= 60 ? 'r-mid' : 'r-bad';
  }

  function renderDashboard() {
    const box = $('#b8s-dash');
    if (!box) return;
    const c = getCalc();
    if (!c) {
      box.innerHTML = '<div class="b8s-empty"><b>Chưa có dữ liệu</b><br>Chọn bưu cục ở thanh phía trên rồi bấm “Chuyển bưu cục”. Script sẽ tự đọc toàn bộ số liệu (tab Tiến độ nhân viên, Chi tiết phát kiện) rồi quay về tab bạn đang xem.</div>';
      return;
    }
    const cl = cardList(c);
    const card = ([label, v, kind, color, sub]) => (v === null || v === undefined || !Number.isFinite(v)) ? ''
      : `<div class="b8s-card"><span>${esc(label)}</span><b class="c-${color}">${esc(fmtVal(v, kind))}</b>${
          Array.isArray(sub) ? sub.map(([k, n]) => `<small>${esc(k)}: <i>${esc(fmt(n))}</i></small>`).join('') : ''}</div>`;
    const cols = shipperCols(c);
    const gh = cols.filter(x => x.group === 'gh').length, tt = cols.filter(x => x.group === 'tt').length;
    const rows = sortRows(c.shippers, state.sort.key, state.sort.dir);
    const th = (x, extra = '') => `<th class="${x.k === 'name' ? 'l ' : ''}g-${x.group || 'nav'} ${extra}" data-k="${x.k}" title="${esc(x.tip || '')}">${esc(x.label)}${state.sort.key === x.k ? (state.sort.dir > 0 ? ' ▲' : ' ▼') : ''}</th>`;
    const td = (x, r) => {
      const v = r[x.k];
      if (x.kind === 'text') return `<td class="l">${esc(v)}</td>`;
      if (x.kind === 'pct') return `<td class="${rateClass(v, x.good)}">${esc(fmtVal(v, 'pct'))}</td>`;
      return `<td>${esc(fmtVal(v, 'int'))}</td>`;
    };
    const head2 = `<tr><th class="g-nav">#</th>${cols.map(x => th(x)).join('')}</tr>`;
    const head1 = `<tr><th class="g-nav" colspan="2">Shipper đi làm: ${fmt(c.nWork)}${c.workersEstimated ? ' (ước tính)' : ''}</th>${gh ? `<th class="g-gh" colspan="${gh}">Chỉ số giao hàng</th>` : ''}${tt ? `<th class="g-tt" colspan="${tt}">Chỉ số TikTok Shop</th>` : ''}</tr>`;
    const body = rows.map((r, i) => `<tr><td class="rk${i < 3 ? i + 1 : ''}">${i + 1}</td>${cols.map(x => td(x, r)).join('')}</tr>`).join('');
    const total = `<tr class="total"><td></td>${cols.map(x => td(x, c.totals)).join('')}</tr>`;

    // Đối chiếu với trang gốc
    const notes = [...c.notes];
    const ov = state.ov;
    if (ov) {
      const chk = (label, a, b) => {
        if (!Number.isFinite(a)) return;
        notes.push(a === b ? `✔ ${label}: ${fmt(b)} khớp trang gốc.` : `⚠ ${label}: tự tính ${fmt(b)} nhưng trang gốc ghi ${fmt(a)}. Kiểm tra lại phân loại trạng thái ở Cài đặt.`);
      };
      chk('Tổng đơn cần phát', ov.need, c.cards.all.total);
      chk('Đơn nhập kho chưa quét phát', ov.unscanned, c.cards.all.unscanned);
    }

    const warn = state.pending ? `<div class="b8s-warn">⚠ Chưa lưu lịch sử: ${esc(state.pending.reason)} <button class="b8sbtn solid" data-act="force-save">Vẫn lưu</button></div>` : '';
    box.innerHTML = `${warn}
      <div class="b8s-sec">📦 Giao thường</div>
      <div class="b8s-cards">${cl.main.map(card).join('')}</div>
      ${c.hasTT ? `<div class="b8s-sec">🎵 TikTok Shop</div><div class="b8s-cards">${cl.tts.map(card).join('')}</div>` : ''}
      <div class="b8s-tw"><table><thead>${head1}${head2}</thead><tbody>${body}${total}</tbody></table></div>
      ${notes.length ? `<div class="b8s-notes">${notes.map(n => `<div>${esc(n)}</div>`).join('')}</div>` : ''}`;
  }

  function renderSettings() {
    renderOffices();
    const box = $('#b8s-map');
    if (!box) return;
    if (!state.orders) { box.innerHTML = '<div class="muted">Chưa có dữ liệu đơn. Chọn bưu cục rồi bấm “Chuyển bưu cục” để liệt kê các trạng thái.</div>'; box.dataset.key = ''; return; }
    const stCol = colBy(state.orders.headers, /trạng thái/i);
    if (stCol < 0) { box.innerHTML = '<div class="muted bad">Không thấy cột “Trạng thái” trong bảng đơn.</div>'; return; }
    const list = statusCounts(state.orders.rows, stCol);
    const sig = JSON.stringify(list) + JSON.stringify(state.statusMap);
    if (box.dataset.key === sig) return;
    box.dataset.key = sig;
    const opt = (v, cur, l) => `<option value="${v}" ${v === cur ? 'selected' : ''}>${l}</option>`;
    box.innerHTML = `<table><thead><tr><th class="l">Trạng thái trong bảng</th><th>Số đơn</th><th class="l">Tính là</th></tr></thead><tbody>${
      list.map(([s, n]) => {
        const cur = classify(s, state.statusMap);
        return `<tr><td class="l">${esc(s)}</td><td>${fmt(n)}</td><td class="l"><select data-status="${esc(s)}">${opt('ok', cur, 'Giao thành công')}${opt('unscanned', cur, 'Chưa quét giao (nhập kho)')}${opt('other', cur, 'Đã quét giao (đang giao, khác)')}</select></td></tr>`;
      }).join('')
    }</tbody></table>`;
  }

  function renderSnapInfo() {
    const box = $('#b8s-snapinfo'); if (!box) return;
    const S = state.snap, dates = [...new Set(S.list.map(r => r.date))].sort(), today = dateKey(new Date()), now = new Date();
    const todayTxt = SNAP_SLOTS.map(sl => { const st = slotState(today, sl, now, OFFICES, S.idx); return `${sl} ${st.st === 'ok' ? '✔' : st.st === 'part' ? `(${st.have}/${st.n})` : st.st === 'miss' ? 'lỡ' : st.st === 'due' ? 'đang chờ' : '–'}`; }).join(' · ');
    box.innerHTML = `Tự lưu vào các mốc <b>${SNAP_SLOTS.join(', ')}</b> (giờ máy, kể cả Chủ nhật) khi trang 800BEST đang mở; mở trễ trong ${SNAP_WINDOW_MIN} phút sau mốc thì chạy bù. Giữ ${SNAP_KEEP_DAYS} ngày, cũ hơn tự xóa.<br>`
      + `Hôm nay: ${esc(todayTxt)}<br>Đang có ${S.list.length} bản chụp${dates.length ? ` (từ ${esc(fmtDMY(dates[0]))} đến ${esc(fmtDMY(dates[dates.length - 1]))})` : ''}.`
      + `${state.fast ? '' : '<br><b class="bad">Đang tắt “Tải nhanh bằng API” nên không tự lưu được.</b>'}${S.err ? `<br><b class="bad">${esc(S.err)}</b>` : ''}`
      + '<br>Mỗi bản chụp có thêm số liệu từng ship (tên, số quét, số giao thành công, giờ quét đầu) để phân tích nguyên nhân; không lưu mã vận đơn hay thông tin người nhận.'
      + '<br>Lưu ý: tool chỉ chạy được khi máy bật, trình duyệt mở và đã đăng nhập 800BEST. Mốc nào lúc đó không mở trang thì bị lỡ và không bù lại được.';
  }

  function renderBkInfo() {
    const box = $('#b8s-bk-st'); if (!box) return;
    const bi = backupInfo();
    box.innerHTML = `Đang có ${state.hist.size} bản ghi ngày và ${state.snap.list.length} bản chụp theo mốc giờ.${bi.due ? ' <b class="bad">Đã hơn 7 ngày chưa sao lưu.</b>' : ''}`;
  }
  function renderUI() {
    if (!$('#b8s-app')) return;
    const title = $('#b8s-title'), sub = $('#b8s-sub');
    $('#b8s-app').classList.toggle('anm', state.view === 'anm');
    if (state.view === 'anm') { if (title) title.textContent = `📍 Thống kê Khu Vực${REGION ? ' ' + REGION : ''}`; }
    else if (title) title.textContent = state.office ? `📍 ${state.office}${state.province ? ' – ' + titleCase(state.province) : ''}` : '📊 Thống kê tổng quan';
    if (state.office && document.activeElement !== $('#b8s-office-in')) renderOfficeList();
    if (sub && state.view !== 'anm') {
      let t = state.lastRefresh ? `Cập nhật lúc ${state.lastRefresh.toLocaleTimeString('vi-VN')} ngày ${state.lastRefresh.toLocaleDateString('vi-VN')}` : 'Chưa tải dữ liệu';
      if (state.lastSaved) t += ` · đã lưu lịch sử ${pad(state.lastSaved.getHours())}:${pad(state.lastSaved.getMinutes())}`;
      if (backupInfo().due) t += ' · ⚠ nên sao lưu lịch sử';
      sub.textContent = t;
    }
    ['dash', 'hist', 'anm', 'set'].forEach(v => {
      const sec = $(`#b8s-v-${v}`); if (sec) sec.style.display = state.view === v ? 'block' : 'none';
    });
    document.querySelectorAll('#b8s-app .b8s-tab').forEach(t => t.classList.toggle('on', t.dataset.view === state.view));
    if (state.view === 'dash') renderDashboard();
    if (state.view === 'hist') renderHistory();
    if (state.view === 'anm') renderAnm();
    if (state.view === 'set') { renderSettings(); renderSnapInfo(); renderWxInfo(); renderBkInfo(); }
    snapBadge();
  }

  /* ------------------------------------------------------------------ *
   * Xuất Excel & chẩn đoán
   * ------------------------------------------------------------------ */
  async function exportExcel() {
    if (state.busy) return;
    if (typeof XLSX === 'undefined') { log('Thư viện Excel chưa tải được. Kiểm tra mạng/CDN rồi tải lại trang.', true); return; }
    if (!state.orders && !state.emp && state.view !== 'hist' && state.view !== 'anm') { if (!(await updateAll())) return; }
    const c = getCalc();
    const now = new Date(), used = new Set(), wb = XLSX.utils.book_new();
    const addSheet = (name, aoa) => {
      const ws = XLSX.utils.aoa_to_sheet(aoa);
      ws['!cols'] = (aoa[0] || []).map((_, i) => ({ wch: Math.min(40, Math.max(10, ...aoa.slice(0, 200).map(r => String(r[i] ?? '').length + 2))) }));
      XLSX.utils.book_append_sheet(wb, ws, sheetName(name, used));
    };
    const val = (v, kind) => (kind === 'text') ? String(v ?? '') : (typeof v === 'number' && Number.isFinite(v)) ? (kind === 'int' ? v : round1(v)) : '';

    if (c) {
      const cl = cardList(c);
      const ov = [['Nhóm', 'Chỉ tiêu', 'Giá trị']];
      const push = (g, list) => list.forEach(([label, v, kind, , sub]) => {
        if (v === null || v === undefined || !Number.isFinite(v)) return;
        ov.push([g, label + (kind === 'pct' ? ' (%)' : ''), kind === 'int' ? v : round1(v)]);
        if (Array.isArray(sub)) sub.forEach(([k, n]) => ov.push([g, `${label} – ${k}`, n]));
      });
      push('Giao thường', cl.main);
      if (c.hasTT) push('TikTok Shop', cl.tts);
      addSheet('Tong quan', ov);

      const cols = shipperCols(c);
      const rows = sortRows(c.shippers, state.sort.key, state.sort.dir);
      addSheet('Shipper', [
        cols.map(x => x.label),
        ...rows.map(r => cols.map(x => val(r[x.k], x.kind))),
        cols.map(x => val(c.totals[x.k], x.kind))
      ]);
    }

    if (state.view === 'anm') addAnmSheets(addSheet);
    addHistorySheets(addSheet, now);

    const notes = [['Nội dung'], [`Bưu cục: ${state.office || 'Bưu cục đang chọn'}${state.province ? ' – ' + titleCase(state.province) : ''}`], [`Thời điểm xuất: ${now.toLocaleString('vi-VN')}`]];
    if (state.lastRefresh) notes.push([`Dữ liệu đọc lúc: ${state.lastRefresh.toLocaleString('vi-VN')}`]);
    notes.push([`Mã vận đơn TikTok Shop: ${state.tiktok}`]);
    notes.push(['POD2400 = đơn mới hôm nay (Ngày cần phát hàng = hôm nay) giao thành công / tổng đơn mới hôm nay, tính cả đơn chưa quét giao. Backlog 3 ngày = đơn có Ngày cần phát cách hôm nay từ 3 ngày trở lên và chưa giao thành công. TikTok đến hạn ODR = đơn TikTok có Thời hạn giao hàng nội bộ là hôm nay.']);
    notes.push(['GTC chuyến = đơn giao thành công / đơn đã quét giao; GTC Total = đơn giao thành công / tổng đơn cần phát. Tổng đơn TTS ở bảng shipper = đơn TikTok đã quét giao.']);
    notes.push(['Tỷ lệ quét giao = (tổng đơn − đơn nhập kho chưa quét) / tổng đơn. TB đơn/NV = tổng đơn / số shipper đi làm hôm nay.']);
    if (c) c.notes.forEach(t => notes.push([t]));
    notes.push([state.includePII ? 'Có kèm thông tin người nhận.' : 'Không kèm thông tin người nhận (SĐT, địa chỉ…).']);
    addSheet('Ghi chu', notes);

    const filename = `800BEST_ThongKe_${safeName(state.office)}_${localStamp(now)}.xlsx`;
    XLSX.writeFile(wb, filename);
    log(`Đã xuất Excel: ${filename}`);
  }

  async function copyText(text) {
    try { await navigator.clipboard.writeText(text); return true; } catch { /* thử cách khác */ }
    const ta = document.createElement('textarea');
    ta.value = text; ta.style.cssText = 'position:fixed;opacity:0';
    document.body.appendChild(ta); ta.select();
    let ok = false;
    try { ok = document.execCommand('copy'); } catch { /* bỏ qua */ }
    ta.remove();
    return ok;
  }

  async function diagnose() {
    const c = getCalc();
    const stCol = state.orders ? colBy(state.orders.headers, /trạng thái/i) : -1;
    const info = {
      version: '3.11.3',
      signTimeKeys: state.sign ? state.sign.keys : null,
      signStats: state.sign ? { doneOrders: state.sign.okN, withSignTime: state.sign.signN } : null,
      weather: { on: state.wx.on, ok: !!state.wx.data, err: state.wx.err || '', hours: state.wx.data ? state.wx.data.times.length : 0 },
      page: location.pathname + location.hash,
      tabs: [...document.querySelectorAll('.el-tabs__item, [role="tab"], .ant-tabs-tab')].filter(visible).map(norm).slice(0, 20),
      employeeHeaders: state.emp ? state.emp.headers : null,
      employeeRows: state.emp ? state.emp.rows.length : 0,
      workCol: c && c.W ? c.W.workHeader : null,
      orderHeaders: state.orders ? state.orders.headers : null,
      orderRows: state.orders ? state.orders.rows.length : 0,
      statuses: state.orders && stCol >= 0 ? statusCounts(state.orders.rows, stCol) : null,
      overview: state.ov || null,
      office: state.office, province: state.province,
      history: { ready: state.histReady, error: state.histErr, records: state.hist.size, offices: [...new Set([...state.hist.values()].map(r => r.office))], lastSaved: state.lastSaved, pending: state.pending ? state.pending.reason : null },
      officeSelector: (() => {
        const i = findOfficeInput();
        return i ? { found: true, value: i.value, shownText: currentOfficeText(i), wrapperClass: selectWrap(i)?.className, placeholder: i.placeholder, itemsBelow: dropdownItems(i).slice(0, 8).map(norm), searchButtons: searchCandidates(i).slice(0, 5).map(x => ({ tag: x.el.tagName, cls: String(x.el.className).slice(0, 60), dy: Math.round(x.dy), dx: Math.round(x.dx) })), readOnly: i.readOnly, inputClass: i.className, wrapClass: i.parentElement?.className, wrapParentClass: i.parentElement?.parentElement?.className, searchButton: !!findSearchBtn(i) } : { found: false };
      })(),
      openDropdownOptions: getOptions().slice(0, 8).map(norm),
      endpoints: Object.fromEntries(state.endpoints)
    };
    const ok = await copyText(JSON.stringify(info, null, 2));
    log(ok ? 'Đã sao chép thông tin chẩn đoán (chỉ gồm tên cột, trạng thái, tên API; không có mã vận đơn hay thông tin người nhận).' : 'Không sao chép được vào clipboard.', !ok);
  }

  /* ------------------------------------------------------------------ *
   * Dựng giao diện
   * ------------------------------------------------------------------ */
  function mount() {
    if ($('#b8s-launch')) return;
    loadPrefs(); loadSites(); loadAi(); loadOffices(); loadDisc();
    const css = `
#b8s-launch.logo{left:12px;top:12px;right:auto;bottom:auto;width:150px;padding:0;background:#fff;border-radius:8px;overflow:hidden;line-height:0;box-shadow:0 3px 12px #0005;border:1px solid #d6dbe4}#b8s-launch.logo .b8sl-t{display:block;line-height:1;padding:6px 0 0;text-align:center;color:#022344;font:800 12.5px/1 system-ui,Segoe UI,Arial,sans-serif;letter-spacing:3px}#b8s-launch.logo img{display:block;width:150px;height:86px}#b8s-launch.logo.bad{outline:3px solid #ef4444;outline-offset:-3px}#b8s-launch.logo:hover{filter:brightness(.96)}
#b8s-launch{position:fixed;right:20px;bottom:22px;z-index:999999;background:#4c3a8f;color:#fff;border:0;border-radius:10px;padding:12px 16px;font-weight:700;cursor:pointer;box-shadow:0 4px 18px #0005}
#b8s-disc{position:absolute;left:50%;transform:translateX(-50%);bottom:14px;z-index:30;width:min(920px,94%);max-height:60%;overflow:auto;box-shadow:0 8px 30px #0006;border-radius:10px;background:#fff}#b8s-disc .al{margin:0}#b8s-disc .b8sbtn.ghost{color:#374151;background:#f3f4f6;border:1px solid #d1d5db}
#b8s-app.lk .lk-hide{display:none}
#b8s-app{position:fixed;inset:2vh 2vw;z-index:1000000;display:none;flex-direction:column;border-radius:16px;overflow:hidden;box-shadow:0 20px 80px #0007;font:14px/1.45 "Segoe UI",Roboto,Arial,sans-serif;background:linear-gradient(135deg,#667eea 0%,#764ba2 100%);color:#1f2937}
#b8s-app *{box-sizing:border-box}
.b8s-top{padding:8px 18px 6px;color:#fff;display:flex;gap:12px;align-items:center;flex-wrap:wrap;justify-content:space-between}
.b8s-top h2{margin:0;font-size:22px;font-weight:800;letter-spacing:.3px;color:#ffe66d;text-shadow:0 2px 6px #2b1b5a99}
.b8s-sub{opacity:.85;font-size:12px;margin-top:2px}
.b8s-acts{display:flex;gap:8px;flex-wrap:wrap}
.b8s-pick{display:flex;gap:8px;flex:1;justify-content:center;min-width:280px}
.b8s-pick select{padding:8px 10px;border:0;border-radius:8px;min-width:220px;max-width:320px;flex:1;font-size:14px;font-weight:600;color:#4c3a8f;background:#fff}
.b8sbtn{border:0;border-radius:8px;background:#fff;color:#4c3a8f;padding:8px 13px;cursor:pointer;font-weight:700}
.b8sbtn.ghost{background:#ffffff33;color:#fff}
.b8s-tabs{display:flex;gap:6px;padding:0 20px}
.b8s-tab{border:0;background:#ffffff2e;color:#fff;padding:6px 16px;border-radius:9px 9px 0 0;cursor:pointer;font-weight:600}
.b8s-tab.on{background:#fff;color:#4c3a8f}
.b8s-main{flex:1;overflow:auto;padding:8px 16px;min-height:0}
.b8s-sec{color:#fff;font-weight:700;text-align:center;margin:8px 0;font-size:15px}
.b8s-cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px}
.b8s-card{background:#fffffff2;border-radius:12px;padding:10px 12px;text-align:center;box-shadow:0 2px 8px #0002}
.b8s-card span{display:block;font-size:12px;color:#4b5563}
.b8s-card b{display:block;font-size:26px;margin-top:2px;line-height:1.2}
.b8s-card small{display:block;font-size:11px;color:#6b7280;line-height:1.5}.b8s-card small i{font-style:normal;font-weight:700;color:#374151}
.c-red{color:#dc2626}.c-blue{color:#2563eb}.c-green{color:#16a34a}.c-amber{color:#d97706}.c-purple{color:#7c3aed}.c-teal{color:#0d9488}.c-orange{color:#ea580c}.c-pink{color:#db2777}
.b8s-tw{background:#fff;border-radius:12px;overflow:auto;margin-top:14px;max-height:60vh;box-shadow:0 2px 8px #0002}
.b8s-tw table,.b8s-panel table{width:100%;border-collapse:collapse;font-size:13px}
.b8s-tw th,.b8s-tw td,.b8s-panel th,.b8s-panel td{padding:8px 10px;text-align:center;border-bottom:1px solid #e5e7eb;white-space:nowrap}
.b8s-tw .l,.b8s-panel .l{text-align:left}
.b8s-tw thead th{position:sticky;color:#fff;cursor:pointer;font-weight:600;line-height:18px;z-index:1}
.b8s-tw thead tr:nth-child(1) th{top:0;height:34px;cursor:default}
.b8s-tw thead tr:nth-child(2) th{top:34px}
.g-nav{background:#1e3a8a}.g-gh{background:#3b82f6}.g-tt{background:#f59e0b}
.b8s-tw tbody tr:nth-child(even){background:#f9fafb}
.b8s-tw tbody tr:hover{background:#eef2ff}
.rk1,.rk2,.rk3{color:#d97706;font-weight:700}
.r-ok{color:#16a34a;font-weight:700}.r-mid{color:#d97706;font-weight:700}.r-bad{color:#dc2626;font-weight:700}
.b8s-tw tr.total td{font-weight:700;background:#eef2f7;border-top:2px solid #c7d2e0;position:sticky;bottom:0}
.b8s-notes{margin-top:12px;color:#fff;font-size:12px;opacity:.95}
.b8s-empty{background:#fffffff2;border-radius:12px;padding:28px;text-align:center;color:#4b5563}
.b8s-panel{background:#fff;border-radius:12px;padding:14px;margin-bottom:12px}
.b8s-setrow{display:grid;grid-template-columns:minmax(240px,1fr) minmax(320px,2fr);gap:12px;align-items:start}@media(max-width:900px){.b8s-setrow{grid-template-columns:1fr}}
.b8s-panel h3{margin:0 0 10px;font-size:15px}
.b8s-panel label{display:block;font-size:12px;color:#4b5563}
.b8s-panel input[type=text],.b8s-panel input[type=number],.b8s-panel select{padding:6px;border:1px solid #cbd5e1;border-radius:6px;font-size:13px}
.row{display:flex;gap:14px;flex-wrap:wrap;align-items:flex-end;margin-bottom:10px}
.muted{color:#6b778c;padding:6px 0;font-size:12px}
#b8s-log{max-height:200px;overflow:auto;font:12px/1.5 monospace;background:#0f172a;color:#dbeafe;padding:8px;border-radius:7px}
.bad{color:#dc2626}#b8s-log .bad{color:#fecaca}
.b8s-warn{background:#fff7ed;color:#9a3412;border:1px solid #fdba74;border-radius:10px;padding:8px 12px;margin-bottom:10px;font-size:13px}
.b8sbtn.solid{background:#4c3a8f;color:#fff}
.b8s-tw small.up,.b8s-tw small.dn{display:block;font-size:10px;font-weight:600}
.b8s-tw small.up{color:#16a34a}.b8s-tw small.dn{color:#dc2626}
.b8sx{border:0;background:none;color:#9ca3af;cursor:pointer;font-size:14px}.b8sx:hover{color:#dc2626}
.b8s-panel input[type=date]{padding:5px;border:1px solid #cbd5e1;border-radius:6px;font-size:13px}
#b8s-hist .b8s-tw{margin-top:0}
#b8s-app .anm-top{display:flex;gap:10px;align-items:center;flex-wrap:wrap;margin-bottom:8px}
#b8s-app .anm-f{color:#fff;font-size:12px;display:flex;gap:6px;align-items:center}
#b8s-app .anm-f input,#b8s-app .anm-f select{padding:6px;border:0;border-radius:8px;font-size:13px;font-weight:600;color:#4c3a8f}
#b8s-app .anm-meta{color:#fff;font-size:12px;margin-bottom:10px;opacity:.95}
#b8s-app .anm-wrap{position:relative}
#b8s-app .btn{display:inline-block;background:#4c3a8f;color:#fff;border-radius:8px;padding:7px 13px;font-size:12px;font-weight:700;cursor:pointer;white-space:nowrap;user-select:none}
#b8s-app .anm-top .btn{background:#fff;color:#4c3a8f}#b8s-app .btn.solid{background:#4c3a8f!important;color:#fff!important;padding:3px 10px;margin-left:8px}
#b8s-app .btn.ghost{background:#ffffff40!important;color:#fff!important}
#b8s-app .b8s-pop{position:absolute;right:0;top:34px;z-index:20;background:#fff;border:1px solid #c4b5fd;border-radius:10px;box-shadow:0 8px 24px #0004;padding:10px;width:330px;color:#1f2937}
#b8s-app .b8s-pop .pg{display:flex;flex-wrap:wrap;gap:5px;margin-bottom:7px}#b8s-app .b8s-pop .pg b{width:100%;font-size:11px;color:#6b7280}
#b8s-app .b8s-pop .o{border:1px solid #cbd5e1;border-radius:12px;padding:3px 10px;font-size:12px;font-weight:600;cursor:pointer}#b8s-app .b8s-pop .o:hover{background:#ede9fe}#b8s-app .b8s-pop .o.on{background:#4c3a8f;color:#fff;border-color:#4c3a8f}
#b8s-app .b8s-pop .hn{font-size:11px;color:#6b7280;border-top:1px solid #e5e7eb;padding-top:6px;margin-top:4px}
#b8s-app .b8s-pop .ol{display:flex;align-items:center;gap:6px;padding:4px 2px;font-size:13px;cursor:pointer;color:#1f2937}
#b8s-app .lnk{cursor:pointer}#b8s-app .lnk:hover{text-decoration:underline}
#b8s-app .dot{display:inline-block;width:9px;height:9px;border-radius:50%;margin-right:6px}
#b8s-app .anm-two{display:grid;grid-template-columns:1.05fr 1fr;gap:12px;margin-top:12px}
@media(max-width:1000px){#b8s-app .anm-two{grid-template-columns:1fr}}
#b8s-app .anm-two .b8s-panel{margin-bottom:0}
#b8s-app .al-box{max-height:230px;overflow:auto}
#b8s-app .anm-hd{display:flex;justify-content:space-between;align-items:center;margin:12px 0 8px}
#b8s-app .anm-two h3{margin:0 0 8px}
#b8s-app .al{padding:6px 9px;border-radius:8px;margin-bottom:5px;font-size:12.5px}#b8s-app .al.r{background:#fee2e2;color:#991b1b}#b8s-app .al.a{background:#fef3c7;color:#92400e}#b8s-app .al.ok{background:#dcfce7;color:#166534}
#b8s-app .rk-row{display:grid;grid-template-columns:22px 150px 1fr 64px 56px;gap:8px;align-items:center;padding:5px 0;font-size:13px}
#b8s-app .rkn{display:inline-block;width:22px;height:22px;line-height:22px;text-align:center;border-radius:11px;background:#e5e7eb;font-weight:700;font-size:12px}
#b8s-app .rkb{height:14px;border-radius:7px;background:#e5e7eb;position:relative;overflow:hidden}#b8s-app .rkb i{position:absolute;left:0;top:0;bottom:0;border-radius:7px}
#b8s-app .rkd{font-size:11px;font-weight:700;text-align:left}
#b8s-app .up{color:#16a34a}#b8s-app .dn{color:#dc2626}#b8s-app .nt{color:#6b7280}
#b8s-app .tdl{display:block;font-size:10px;font-weight:700}
#b8s-app .b8s-card small.up{color:#16a34a}#b8s-app .b8s-card small.dn{color:#dc2626}
#b8s-app .chips{display:flex;gap:5px;flex-wrap:wrap;margin:6px 0}
#b8s-app .ch{border:1.5px solid var(--c,#4c3a8f);color:var(--c,#4c3a8f);border-radius:12px;padding:2px 9px;font-size:11px;font-weight:600;cursor:pointer;user-select:none}
#b8s-app .ch.off{opacity:.35;text-decoration:line-through}#b8s-app .ch.all{background:#4c3a8f;color:#fff;--c:#4c3a8f}
#b8s-app .b8s-tw td.h1{background:#dcfce7;color:#166534;font-weight:700}#b8s-app .b8s-tw td.h2{background:#fef3c7;color:#92400e;font-weight:700}#b8s-app .b8s-tw td.h3{background:#fee2e2;color:#991b1b;font-weight:700}
#b8s-anm .b8s-cards{gap:8px;grid-template-columns:repeat(6,1fr)}
#b8s-anm .anm-two{grid-template-columns:5fr 7fr}
#b8s-anm .anm-bot{display:grid;grid-template-columns:3fr 2fr;gap:8px;margin-top:8px;align-items:start}
#b8s-anm .anm-bot .b8s-tw{margin-top:0}
#b8s-anm .anm-alert{max-height:300px;overflow:auto}
@media(max-width:1000px){#b8s-anm .anm-bot,#b8s-anm .anm-two{grid-template-columns:1fr}}
#b8s-anm .b8s-card{padding:5px 8px;border-radius:10px}
#b8s-anm .b8s-card span{font-size:11px}
#b8s-anm .b8s-card b{font-size:22px;margin-top:0;line-height:1.15}
#b8s-anm .b8s-card small{font-size:10px;line-height:1.3}
#b8s-anm .anm-top{margin-bottom:4px;gap:8px}
#b8s-anm .anm-meta{margin-bottom:6px;font-size:11px}
#b8s-anm .anm-two{gap:8px;margin-top:8px}
#b8s-anm .b8s-panel{padding:8px 10px;border-radius:10px}
#b8s-anm .anm-two h3{font-size:13px;margin:0 0 4px}
#b8s-anm .anm-hd{margin:8px 0 4px}
#b8s-anm .al{padding:3px 8px;font-size:12.5px;margin-bottom:3px;line-height:1.35}
#b8s-anm .al-box{max-height:none}
#b8s-anm .rk-row{padding:2px 0;font-size:13px;grid-template-columns:20px 130px 1fr 58px 52px}
#b8s-anm .rkn{width:18px;height:18px;line-height:18px;font-size:11px}
#b8s-anm .rkb{height:11px}
#b8s-anm .chips{margin:3px 0}
#b8s-anm .ch{padding:1px 8px;font-size:11px}
#b8s-anm .muted{padding:2px 0;font-size:11px}
#b8s-anm .b8s-tw{margin-top:8px;max-height:none;overflow:visible}
#b8s-anm .b8s-tw table{font-size:13px}
#b8s-anm .b8s-tw th,#b8s-anm .b8s-tw td{padding:3px 6px}
#b8s-anm .b8s-tw thead th{position:static;line-height:15px}
#b8s-anm .b8s-tw tr.total td{position:static}
#b8s-anm .tdl{display:inline;font-size:9px;margin-left:4px}
#b8s-anm .b8s-notes{margin-top:6px;font-size:11px}
#b8s-anm .seg{display:inline-flex;background:#ede9fe;border-radius:8px;padding:2px}
#b8s-anm .seg span{padding:2px 9px;border-radius:6px;font-size:11px;font-weight:700;color:#5b21b6;cursor:pointer;user-select:none}
#b8s-anm .seg span.on{background:#4c3a8f;color:#fff}
#b8s-anm .rk-row.two{grid-template-columns:20px 120px 1fr 54px 108px 108px}
#b8s-anm .rk-row.two .rkb{overflow:visible}
#b8s-anm .rkh{display:grid;grid-template-columns:20px 120px 1fr 54px 108px 108px;gap:8px;font-size:10px;color:#64748b;font-weight:700;padding:0 0 2px}
#b8s-anm .mk{position:absolute;top:-3px;bottom:-3px;width:3px;border-radius:2px;background:#1f2937}
#b8s-anm .mk.w{background:#6366f1}
#b8s-anm .cd{font-size:11px;font-weight:700;white-space:nowrap}#b8s-anm .cd small{font-weight:500;color:#64748b;font-size:10px;margin-left:3px}
#b8s-anm .mv{font-size:10px;font-weight:700}
#b8s-anm .lg{display:inline-block;height:11px;border-left:3px solid;margin-right:4px;vertical-align:middle}
#b8s-anm .slots{display:flex;gap:5px;align-items:center;color:#fff;font-size:11px;margin:2px 0 6px;flex-wrap:wrap}
#b8s-anm .slot{padding:2px 9px;border-radius:12px;background:#ffffff33;font-weight:700;color:#fff;cursor:pointer;user-select:none}
#b8s-anm .slot.fut{opacity:.5;cursor:default}
#b8s-anm .slot.miss{background:#fecaca;color:#b91c1c;text-decoration:line-through;cursor:default}
#b8s-anm .slot.due{background:#fde68a;color:#92400e;cursor:default}
#b8s-anm .slot.part{background:#fef3c7;color:#92400e}
#b8s-anm .slot.live{background:#bbf7d0;color:#166534}
#b8s-anm .slot.on{background:#fff;color:#4c3a8f}
#b8s-anm .slot.live.on{background:#bbf7d0;color:#166534;outline:2px solid #fff}
#b8s-anm .b8s-card .cmp2{display:flex;justify-content:center;gap:10px;font-size:10px;color:#64748b;font-weight:600}
#b8s-anm .b8s-card .cmp2 small{display:inline;font-size:10px;font-weight:700}
#b8s-anm .rkp{margin-top:8px;overflow-x:auto}
#b8s-anm .rk2{display:grid;align-items:center;gap:8px;padding:2px 6px;border-radius:9px;white-space:nowrap;font-size:13px}
#b8s-anm .rk2:nth-child(even){background:#f8fafc}
#b8s-anm .rk2.hd{font-size:10.5px;color:#64748b;font-weight:700;background:none;padding-bottom:2px}
#b8s-anm .rk2 .rkb{height:11px;overflow:visible}
#b8s-anm .rk2 .rkt{font-weight:700;overflow:hidden;text-overflow:ellipsis}
#b8s-anm .med{display:inline-block;width:24px;height:24px;border-radius:12px;text-align:center;line-height:24px;font-weight:800;font-size:12px;color:#fff;background:#94a3b8}
#b8s-anm .med.m1{background:linear-gradient(135deg,#fbbf24,#d97706)}#b8s-anm .med.m2{background:linear-gradient(135deg,#cbd5e1,#64748b)}#b8s-anm .med.m3{background:linear-gradient(135deg,#fdba74,#b45309)}
#b8s-anm .pill{display:inline-block;border-radius:6px;padding:1px 8px;text-align:center}
#b8s-anm .pill.h1{background:#dcfce7;color:#166534}#b8s-anm .pill.h2{background:#fef3c7;color:#92400e}#b8s-anm .pill.h3{background:#fee2e2;color:#991b1b}
#b8s-anm .toplin{display:flex;justify-content:center;align-items:center;gap:10px;flex-wrap:wrap;padding:3px 8px 5px;font-size:13px;color:#1f2937}
#b8s-anm .toplin em{font-style:normal;color:#cbd5e1}#b8s-anm .toplin i{font-style:normal;font-weight:800}#b8s-anm .toplin b{color:#b45309}
#b8s-anm .toplin .lnk{font-weight:800;cursor:pointer}
#b8s-anm .anm-bot2 .anm-an{max-height:230px}
#b8s-anm.tiny .anm-meta,#b8s-anm.tiny .chips,#b8s-anm.tiny details.b8s-notes{display:none}
#b8s-anm.tiny .anm-bot2 .anm-an{max-height:140px}#b8s-anm.tiny .anm-bot2 svg{height:105px}
#b8s-anm.tiny .b8s-card{padding-top:4px;padding-bottom:4px}
#b8s-anm .b8s-notes>summary{cursor:pointer;font-weight:700;color:#fff;font-size:11px}
#b8s-anm details.b8s-notes{padding:4px 8px}
#b8s-app.anm .b8s-pick,#b8s-app.anm #b8s-export,#b8s-app.anm .b8s-sub{display:none}
#b8s-app.anm .b8s-top>div:first-child{flex:1;text-align:center;padding-left:96px}
#b8s-anm .rk2.hd{white-space:normal;line-height:1.2;align-items:end}
#b8s-anm .qs{margin-left:4px;font-size:10px;font-weight:500;color:#64748b}
#b8s-anm .fc b{font-size:14px;color:#111827}#b8s-anm .fc.base b{color:#6b7280;font-style:italic}#b8s-anm .fcs{margin-left:4px;font-size:10px;font-weight:500;color:#64748b}
#b8s-anm .aibar{display:flex;align-items:center;gap:10px;flex-wrap:wrap;justify-content:flex-end;padding:8px 10px 2px;margin-top:4px;border-top:1px solid #eef2f7;font-size:12px}
#b8s-anm .sw{display:inline-block;width:38px;height:20px;border-radius:10px;background:#cbd5e1;position:relative;cursor:pointer;flex:none}
#b8s-anm .sw i{position:absolute;top:2px;left:2px;width:16px;height:16px;border-radius:50%;background:#fff;transition:left .15s}
#b8s-anm .sw.on{background:#4c3a8f}#b8s-anm .sw.on i{left:20px}#b8s-anm .sw.dis{opacity:.55}
#b8s-anm .anm-bot2{display:grid;grid-template-columns:2.5fr 1fr;gap:8px;margin-top:8px;align-items:stretch;min-height:var(--botmin,0)}
#b8s-anm .cod .cdh{display:flex;align-items:center;gap:8px;flex-wrap:wrap}
#b8s-anm .cdchip{display:inline-block;border-radius:9px;padding:1px 10px;font-size:12px;font-weight:700;background:#f1f5f9;color:#64748b}
#b8s-anm .cdchip.on{background:#ede9fe;color:#4c3aa8}#b8s-anm .cdchip.rd{background:#fee2e2;color:#991b1b}
#b8s-anm .cdcols{display:grid;grid-template-columns:1.5fr 1fr;gap:14px;margin-top:6px;flex:1;min-height:0;align-content:start}
#b8s-anm .cdt{display:grid;grid-template-columns:minmax(96px,1.3fr) minmax(98px,1fr) minmax(98px,1fr) minmax(98px,1fr) 56px 46px;gap:6px;align-items:center;padding:var(--cdpad,5px) 6px;font-size:14.5px;white-space:nowrap}
#b8s-anm .cdzero{margin-top:6px;border:1px solid #fecaca;background:#fff7f7;border-radius:8px;padding:6px 8px}
#b8s-anm .cdzl{max-height:150px;overflow:auto}
#b8s-anm .zc{cursor:pointer;text-decoration:underline;padding:0 3px;border-radius:4px}
#b8s-anm .zc.on{background:#fecaca}
#b8s-anm .cdt.hd{font-size:11.5px;color:#64748b;font-weight:700}
#b8s-anm .cdt:not(.hd):nth-child(even){background:#f8fafc}
#b8s-anm .cdt.tot{border-top:2px solid #e2e8f0;background:#f1f5f9!important;font-weight:800}
#b8s-anm .cdt .nm{font-weight:700;overflow:hidden;text-overflow:ellipsis}
#b8s-anm .cdt .rd,#b8s-anm .cdr .rd{color:#dc2626}#b8s-anm .cdt .gy{color:#94a3b8}
#b8s-anm .cdb{height:8px;border-radius:5px;background:#e5e7eb;position:relative;display:block}#b8s-anm .cdb i{position:absolute;left:0;top:0;bottom:0;border-radius:5px}
#b8s-anm .cdal{max-height:var(--cdal,215px);overflow:auto}#b8s-anm .cod{display:flex;flex-direction:column}
#b8s-anm .cda{display:grid;grid-template-columns:11px minmax(0,1fr) auto;gap:7px;align-items:center;padding:3px 6px;font-size:13.5px}
#b8s-anm .cda:nth-child(even){background:#fff7f7}
#b8s-anm .cda small{display:block;color:#64748b;font-size:11.5px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
#b8s-anm .cdn{display:block;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
#b8s-anm .cdr{text-align:right;white-space:nowrap}
#b8s-anm .cdd{width:11px;height:11px;border-radius:50%;background:#dc2626}#b8s-anm .cdd.o{background:#f97316}#b8s-anm .cdd.y{background:#eab308}
#b8s-anm details.cdchk table{width:100%;border-collapse:collapse;font-size:11.5px;background:#fff}#b8s-anm details.cdchk th{color:#475569;background:#f1f5f9;padding:2px 6px;text-align:right;font-weight:700;font-size:10.5px}#b8s-anm details.cdchk td{padding:2px 6px;text-align:right;border-top:1px solid #eef2f7;color:#1f2937}#b8s-anm details.cdchk th.l,#b8s-anm details.cdchk td.l{text-align:left}
#b8s-anm details.cdchk{margin-top:2px;font-size:12px}#b8s-anm .cod{padding-bottom:6px}#b8s-anm .cod h3{font-size:16px}#b8s-anm .cod .cdh b{font-size:14.5px!important}#b8s-anm details.cdchk>summary{cursor:pointer;font-weight:700;color:#4c3a8f;font-size:11px}
#b8s-anm.tiny .cdal{max-height:84px}#b8s-anm.tiny details.cdchk{display:none}#b8s-anm.tiny .cdt{padding:0 5px;font-size:11px}#b8s-anm.tiny .cod{padding-top:4px;padding-bottom:4px}#b8s-anm.tiny .cdcols{margin-top:2px}#b8s-anm.tiny .cod small.muted{display:none}
@media(max-width:1000px){#b8s-anm .cdcols{grid-template-columns:1fr}}
@media(max-width:1000px){#b8s-anm .anm-bot2{grid-template-columns:1fr}}
#b8s-anm .anm-an{max-height:420px;overflow:auto}
#b8s-anm .rsn{padding:6px 9px;border-radius:10px;margin-bottom:5px;cursor:pointer;border-left:5px solid;font-size:12.5px}
#b8s-anm .rsn.u{background:#f0fdf4;border-color:#16a34a}#b8s-anm .rsn.d{background:#fef2f2;border-color:#dc2626}
#b8s-anm .rsn .go{float:right;font-size:11px;color:#4c3a8f;font-weight:700}
#b8s-anm .rc{display:inline-block;font-size:11.5px;background:#fff;border:1px solid #e5e7eb;border-radius:10px;padding:1px 8px;margin:3px 4px 0 0;white-space:normal}
#b8s-anm .rsf{font-size:12px;color:#475569;padding:3px 0}
#b8s-anm .tiles{display:flex;gap:8px}
#b8s-anm .tile{flex:1;background:#f8fafc;border-radius:10px;padding:6px;text-align:center}
#b8s-anm .tile.off{opacity:.55}#b8s-anm .tile .tl{font-size:11px;color:#64748b}#b8s-anm .tile b{display:block;font-size:22px;line-height:1.2}#b8s-anm .tile small{font-size:11px;color:#475569}
#b8s-status{padding:6px 20px;color:#fff;font-size:12px;min-height:26px;background:#0000002e}
#b8s-status.bad{color:#fecaca}
@media(max-width:800px){#b8s-app{inset:1vh 1vw}.b8s-card b{font-size:21px}}`;
    const style = document.createElement('style');
    style.id = 'b8s-style';
    style.textContent = css;
    document.documentElement.appendChild(style);

    const launch = document.createElement('button');
    launch.id = 'b8s-launch';
    if (LOGO) { launch.classList.add('logo'); launch.innerHTML = `<span class="b8sl-t">THỐNG KÊ</span><img alt="Thống kê" src="${LOGO}">`; launch.title = 'Thống kê 800BEST'; }
    else launch.textContent = '📊 Thống kê 800BEST';
    document.documentElement.appendChild(launch);

    const app = document.createElement('div');
    app.id = 'b8s-app';
    app.innerHTML = `
<div class="b8s-top">
  <div><h2 id="b8s-title">📊 Thống kê tổng quan</h2><div class="b8s-sub" id="b8s-sub">Chưa cập nhật</div></div>
  <div class="b8s-pick">
    <select id="b8s-office-in"></select>
    <button class="b8sbtn" data-act="switch" id="b8s-switch" title="Chuyển sang bưu cục đã chọn và tự tải toàn bộ số liệu. Chọn lại đúng bưu cục đang xem để tải lại.">Chuyển bưu cục</button>
  </div>
  <div class="b8s-acts">
    <button class="b8sbtn" data-act="export" id="b8s-export">⬇ Xuất Excel</button>
    <button class="b8sbtn ghost" id="b8s-close">Đóng ×</button>
  </div>
</div>
<div class="b8s-tabs">
  <button class="b8s-tab on" data-view="dash">Tổng quan</button>
  <button class="b8s-tab" data-view="anm">ANM</button>
  <button class="b8s-tab" data-view="set">Cài đặt</button>
</div>
<div id="b8s-disc" style="display:none"></div>
<div class="b8s-main">
  <section id="b8s-v-dash"><div id="b8s-dash"></div></section>
  <section id="b8s-v-hist" style="display:none"><div id="b8s-hist"></div></section>
  <section id="b8s-v-anm" style="display:none"><div id="b8s-anm"></div></section>
  <section id="b8s-v-set" style="display:none">
    <div class="b8s-setrow">
    <div class="b8s-panel">
      <h3>TikTok Shop</h3>
      <div class="row"><label>Mã đầu vận đơn của TikTok Shop (cách nhau dấu phẩy)<br><input type="text" id="b8s-tt" style="width:260px"></label></div>
    </div>
    <div class="b8s-panel">
      <h3>🧑‍✈️ Ship đi làm</h3>
      <p class="muted">Chỉ ship đủ điều kiện mới được tính vào “DM làm / Đơn mỗi DM” và hiện trong bảng shipper (ở cả tab Tổng quan và tab ANM). Đơn của ship bị ẩn vẫn nằm trong các số tổng.</p>
      <div class="row"><label><input type="checkbox" id="b8s-dm-on"> Bật lọc ship đi làm</label></div>
      <div class="row"><label>Chức vụ được tính (mỗi dòng một chức vụ, đúng như cột “Chức vụ”)<br><textarea id="b8s-dm-depts" rows="4" style="width:360px"></textarea></label></div>
      <div class="row"><label>Quét đơn đầu tiên muộn hơn giờ này thì không tính <input type="text" id="b8s-dm-cut" style="width:70px" placeholder="14:00"></label></div>
      <div class="row"><label>Cuối ngày (từ 23:30 và khi xem cả ngày): giao thành công dưới <input type="number" id="b8s-dm-min" min="0" style="width:60px"> đơn thì không tính</label></div>
    </div>
    <div class="b8s-panel">
      <h3>🏢 Bưu cục quản lý</h3>
      <div class="muted">Tool tự tìm bưu cục khi bạn bấm tải số liệu. Muốn thêm thủ công: gõ đúng tên bưu cục (đúng cả chữ hoa, chữ thường và dấu) rồi bấm “Thêm”. Tool kiểm tra tên và chạy thử lấy đơn; bưu cục không tồn tại hoặc không có đơn nào (rỗng) sẽ không được thêm. “Mã COD” là mã bưu cục trên trang quản lý COD (trường signSiteId của yêu cầu listCodPayment), không có thì bảng COD của bưu cục đó để trống.</div>
      <div class="row"><label style="display:flex;gap:6px;align-items:center"><input type="checkbox" id="b8s-disc-on"> Tự tìm bưu cục đang quản lý (bưu cục có đơn hôm nay) khi bấm “Tải số liệu” ở ANM hoặc “Chuyển bưu cục” ở Tổng quan. Không bao giờ tự xóa bưu cục.</label></div>
      <div class="row"><input type="text" id="b8s-of-name" autocomplete="off" spellcheck="false" placeholder="Tên bưu cục, ví dụ: tên đúng như trong ô chọn Bưu cục của trang" style="width:260px"><button class="b8sbtn solid" id="b8s-of-add">Thêm</button></div>
      <div id="b8s-of-st" class="muted" style="min-height:18px"></div>
      <div id="b8s-of-list"></div>
    </div>
    </div>
    <div class="b8s-panel">
      <h3>Phân loại trạng thái đơn</h3>
      <p class="muted">Chọn mỗi trạng thái được tính là gì. “Chưa quét giao” là đơn đã nhập kho nhưng chưa quét phát; mọi đơn khác (đang giao, thất bại, hoàn…) là “Đã quét giao”. Lựa chọn được lưu cho lần sau. Các dòng ✔/⚠ dưới bảng ở trang Tổng quan cho biết số tự tính có khớp trang gốc không.</p>
      <div id="b8s-map" style="overflow:auto;max-height:40vh"></div>
    </div>
    <div class="b8s-panel lk-hide">
      <h3>Ngưỡng màu và cảnh báo (tab ANM)</h3>
      <div id="b8s-th" style="overflow:auto"></div>
    </div>
    <div class="b8s-panel lk-hide">
      <h3>Thời tiết (phân tích nguyên nhân POD2400)</h3>
      <div class="row"><label style="display:flex;gap:6px;align-items:center"><input type="checkbox" id="b8s-wx-on"> Dùng lượng mưa theo giờ của Open-Meteo (miễn phí, không cần khóa)</label></div>
      <div class="row"><label>Vĩ độ<br><input type="number" step="any" id="b8s-wx-lat" style="width:120px"></label><label>Kinh độ<br><input type="number" step="any" id="b8s-wx-lon" style="width:120px"></label><button class="b8sbtn solid" id="b8s-wx-test">Tải thử</button></div>
      <div class="muted">Bản này tắt sẵn và để tọa độ Hà Nội; hãy nhập tọa độ nơi bạn làm việc (các bưu cục gần nhau dùng chung một điểm) rồi tích bật. Chỉ gửi tọa độ ra ngoài, không gửi số liệu bưu cục.</div>
      <div id="b8s-wx-st" class="muted"></div>
    </div>
    <div class="b8s-panel lk-hide">
      <h3>🤖 AI dự báo POD cuối ngày</h3>
      <div class="muted">Nhập khóa API của AI rồi bấm OK. Khóa đúng sẽ báo “Thành công”; khi đó công tắc “Dự báo POD bằng AI” dưới bảng xếp hạng (tab ANM) dùng được. Chỉ gửi các con số tổng hợp của bưu cục, không gửi mã đơn, tên ship hay thông tin người nhận. Mỗi lần phân tích tốn một lần gọi AI (tính phí trên tài khoản API của anh). Khóa được lưu trong trình duyệt này; chỉ dùng khóa riêng cho việc này, đặt hạn mức chi tiêu thấp.</div>
      <div class="row"><input type="password" id="b8s-ai-key" autocomplete="off" spellcheck="false" style="width:320px" placeholder="Dán khóa API vào đây"><label style="display:flex;gap:6px;align-items:center">Mô hình <input type="text" id="b8s-ai-model" style="width:150px"></label><button class="b8sbtn solid" id="b8s-ai-ok">OK</button><button class="b8sbtn" id="b8s-ai-clear">Xóa khóa</button></div>
      <div id="b8s-ai-st"></div>
      <div style="margin-top:8px;font-weight:700">Hướng dẫn cho AI (có sẵn mặc định, sửa được)</div>
      <textarea id="b8s-ai-prompt" rows="7" style="width:100%;box-sizing:border-box;font-size:12px"></textarea>
      <div class="row"><button class="b8sbtn solid" id="b8s-ai-psave">Lưu hướng dẫn</button><button class="b8sbtn" id="b8s-ai-preset">Khôi phục mặc định</button><span id="b8s-ai-pst" class="muted"></span></div>
    </div>
    <div class="b8s-panel">
      <h3>Sao lưu / khôi phục lịch sử</h3>
      <div class="muted">Tab Lịch sử đã bỏ, số liệu vẫn được lưu như cũ (bản ghi ngày và bản chụp theo mốc giờ). Sao lưu ra file để phòng khi xóa dữ liệu trình duyệt.</div>
      <div class="row"><button class="b8sbtn solid" id="b8s-bk">Sao lưu</button><button class="b8sbtn solid" id="b8s-rs">Khôi phục</button><input type="file" id="b8s-rs-file" accept=".json,application/json" style="display:none"><button class="b8sbtn" id="b8s-gohist" title="Bảng lịch sử theo kỳ, shipper tổng hợp, so sánh các bưu cục">Mở bảng lịch sử chi tiết (nâng cao)</button></div>
      <div id="b8s-bk-st" class="muted"></div>
    </div>
    <div class="b8s-panel lk-hide">
      <h3>Bản chụp theo mốc giờ (tab ANM)</h3>
      <div id="b8s-snapinfo" class="muted"></div>
    </div>
    <div class="b8s-panel">
      <h3>Công cụ <small id="b8s-ver" style="font-weight:400;opacity:.7"></small></h3>
      <div class="row"><label style="display:flex;gap:6px;align-items:center"><input type="checkbox" id="b8s-fast"> Tải nhanh bằng API khi bấm “Tải số liệu” ở tab ANM (lỗi sẽ tự dùng cách thường)</label></div>
      <div class="row"><button class="b8sbtn" style="background:#4c3a8f;color:#fff" id="b8s-diag">Chẩn đoán API</button></div>
      <div id="b8s-log"></div>
    </div>
  </section>
</div>
<div id="b8s-status"></div>`;
    document.documentElement.appendChild(app);

    launch.onclick = () => { app.style.display = 'flex'; readOverview(); renderUI(); onEnterAnm(); };
    $('#b8s-close').onclick = () => { app.style.display = 'none'; };
    $('#b8s-export').onclick = exportExcel;
    $('#b8s-diag').onclick = diagnose;
    {
      const on = $('#b8s-wx-on'), lat = $('#b8s-wx-lat'), lon = $('#b8s-wx-lon');
      on.checked = state.wx.on; lat.value = state.wx.lat; lon.value = state.wx.lon;
      on.onchange = () => { state.wx.on = on.checked; savePrefs(); renderWxInfo(); if (on.checked) loadWeather(true); };
      const coord = () => {
        const a = parseFloat(lat.value), b = parseFloat(lon.value);
        if (!Number.isFinite(a) || Math.abs(a) > 90 || !Number.isFinite(b) || Math.abs(b) > 180) { lat.value = state.wx.lat; lon.value = state.wx.lon; log('Tọa độ không hợp lệ (vĩ độ −90..90, kinh độ −180..180).', true); return; }
        state.wx.lat = a; state.wx.lon = b; state.wx.data = null; savePrefs(); if (state.wx.on) loadWeather(true); renderWxInfo();
      };
      lat.onchange = coord; lon.onchange = coord;
      { const nm = $('#b8s-of-name'), ls = $('#b8s-of-list');
        $('#b8s-of-add').onclick = officesAdd; nm.addEventListener('keydown', e => { if (e.key === 'Enter') officesAdd(); });
        ls.addEventListener('click', e => { const b = e.target.closest && e.target.closest('[data-del]'); if (b && !b.disabled) officesRemove(b.dataset.del); });
        ls.addEventListener('change', e => { const i = e.target.closest && e.target.closest('input[data-cod]'); if (i) officesSaveCod(i.dataset.cod, i.value); }); }
      $('#b8s-wx-test').onclick = () => { if (!state.wx.on) { state.wx.on = true; on.checked = true; savePrefs(); } loadWeather(true); renderWxInfo(); };
      { // AI
        const mi = $('#b8s-ai-model'), ta = $('#b8s-ai-prompt'), pst = $('#b8s-ai-pst');
        mi.value = state.ai.model; ta.value = state.ai.prompt || AI_PROMPT_DEFAULT;
        $('#b8s-ai-ok').onclick = ((...a) => aiSubmitKey(...a));
        $('#b8s-ai-key').addEventListener('keydown', e => { if (e.key === 'Enter') aiSubmitKey(); });
        $('#b8s-ai-clear').onclick = ((...a) => aiClearKey(...a));
        $('#b8s-ai-psave').onclick = () => { const v = ta.value.trim().slice(0, 4000); state.ai.prompt = v === AI_PROMPT_DEFAULT ? '' : v; saveAi(); pst.textContent = 'Đã lưu hướng dẫn.'; };
        $('#b8s-ai-preset').onclick = () => { ta.value = AI_PROMPT_DEFAULT; state.ai.prompt = ''; saveAi(); pst.textContent = 'Đã khôi phục mặc định.'; };
        renderAiInfo();
      }
      $('#b8s-bk').onclick = backupHistory;
      $('#b8s-rs').onclick = () => $('#b8s-rs-file').click();
      $('#b8s-rs-file').onchange = e => { const f = e.target.files && e.target.files[0]; e.target.value = ''; if (f) restoreHistory(f); };
      $('#b8s-gohist').onclick = () => { state.view = 'hist'; renderUI(); };
    }
    { const v = $('#b8s-ver'); if (v) { const upd = () => { v.textContent = `v3.14.0 · tiêu đề API: ${Object.keys(state.apiHdr).length ? 'đã bắt từ trang' : 'dùng mặc định'}${window.__b8sn > 1 ? ` · ⚠ đang chạy ${window.__b8sn} bản script cùng lúc, hãy tắt bớt trong Tampermonkey` : ''}`; }; upd(); setInterval(upd, 1500); } }
    { const dc = $('#b8s-disc-on'); if (dc) { dc.checked = !!disc.on; dc.onchange = () => { disc.on = dc.checked; saveDisc(); }; } }
    { const fa = $('#b8s-fast'); if (fa) { fa.checked = !!state.fast; fa.onchange = () => { state.fast = fa.checked; savePrefs(); }; } }
    const doSwitch = async () => {
      if (state.busy || disc.busy) return;
      const chosen = $('#b8s-office-in').value;   // giữ lựa chọn của người dùng dù ô chọn có được vẽ lại
      await discoverOffices();   // tự dò bưu cục mới (chỉ gọi nhỏ nếu không có gì mới)
      if (disc.pick) return;
      const sel = $('#b8s-office-in'); if (chosen && sel && OFFICES.includes(chosen)) sel.value = chosen;
      const n = chosen && OFFICES.includes(chosen) ? chosen : $('#b8s-office-in').value;
      if (n) switchOffice(n); else log(OFFICES.length ? 'Hãy chọn một bưu cục.' : 'Chưa có bưu cục nào có đơn hôm nay để chọn.', !!OFFICES.length);
    };
    $('#b8s-disc').addEventListener('click', e => {
      const b = e.target.closest('[data-d]'); if (!b) return;
      const a = b.dataset.d;
      if (a === 'undo') discUndo();
      else if (a === 'close') { disc.note = ''; disc.undo = null; discRender(); }
      else if (a === 'padd') discPickDone(true);
      else if (a === 'pskip') discPickDone(false);
      else if (a === 'pall' || a === 'pnone') { (disc.pick || []).forEach(c => { c.on = a === 'pall'; }); const h = $('#b8s-disc'); if (h) h.dataset.h = ''; discRender(); }
    });
    $('#b8s-disc').addEventListener('change', e => {
      const c = e.target.closest('input[data-dpick]'); if (!c || !disc.pick) return;
      const it = disc.pick[+c.dataset.dpick]; if (it) it.on = c.checked;
    });
    $('#b8s-switch').onclick = doSwitch;
    app.querySelectorAll('.b8s-tab').forEach(t => t.addEventListener('click', () => { state.view = t.dataset.view; renderUI(); onEnterAnm(); }));

    $('#b8s-dash').addEventListener('click', e => {
      if (e.target.closest('[data-act="force-save"]')) { forceSave(); return; }
      const th = e.target.closest('th[data-k]');
      if (!th) return;
      const key = th.dataset.k;
      state.sort = state.sort.key === key ? { key, dir: -state.sort.dir } : { key, dir: key === 'name' ? 1 : -1 };
      renderDashboard();
    });

    $('#b8s-tt').value = state.tiktok;
    const dmFill = () => { const r = dmRulesOf(state.dm); $('#b8s-dm-on').checked = r.on; $('#b8s-dm-depts').value = r.depts.join('\n'); $('#b8s-dm-cut').value = r.cutoff; $('#b8s-dm-min').value = r.minOk; };
    const dmSave = () => {
      state.dm = dmRulesOf({ on: $('#b8s-dm-on').checked, depts: $('#b8s-dm-depts').value.split('\n'), cutoff: $('#b8s-dm-cut').value, minOk: $('#b8s-dm-min').value });
      dmFill(); savePrefs(); state.calc = null; state.calcKey = ''; renderUI();
    };
    dmFill();
    for (const id of ['on', 'depts', 'cut', 'min']) $('#b8s-dm-' + id).addEventListener('change', dmSave);
    $('#b8s-tt').addEventListener('change', e => { state.tiktok = e.target.value.trim() || 'TTVN'; e.target.value = state.tiktok; savePrefs(); renderUI(); });
    $('#b8s-map').addEventListener('change', e => {
      const sel = e.target.closest('select[data-status]');
      if (!sel) return;
      state.statusMap[sel.dataset.status] = sel.value;
      savePrefs(); renderUI();
    });

    let rz = 0; window.addEventListener('resize', () => { clearTimeout(rz); rz = setTimeout(() => { if (state.view === 'anm') fitAnm(); }, 150); });
    $('#b8s-anm').addEventListener('click', ((...a) => onAnmClick(...a)));
    $('#b8s-anm').addEventListener('change', ((...a) => onAnmChange(...a)));
    $('#b8s-th').addEventListener('change', e => {
      const t = e.target, f = t.dataset && t.dataset.th;
      if (!f) return;
      const [k, key] = f.split('.'), v = parseFloat(t.value);
      if (!Number.isFinite(v)) { renderThresholds(); return; }
      state.th[k] = Object.assign({}, state.th[k], { [key]: v });
      savePrefs(); renderUI();
    });
    $('#b8s-th').addEventListener('click', e => { if (e.target.id === 'b8s-th-reset') { state.th = {}; savePrefs(); renderThresholds(); renderUI(); } });
    renderThresholds();
    $('#b8s-hist').addEventListener('change', onHistChange);
    $('#b8s-hist').addEventListener('click', onHistClick);
    renderOfficeList();
    renderUI();
    loadHistory();
    loadSnapLog();
    loadSnaps().then(() => { renderUI(); snapBadge(); });
    if (!state.snap.timer) { state.snap.timer = setInterval(((...a) => snapTick(...a)), 30000); setTimeout(((...a) => snapTick(...a)), 8000); }
    lkBoot();
  }

  // Khởi động an toàn: lỗi nào cũng hiện nút báo lỗi (thay vì im lặng), và tự dựng lại nếu trang web xóa mất nút
  function showFatal(e) {
    try {
      console.error('[800BEST]', e);
      if (document.getElementById('b8s-fatal') || !document.documentElement) return;
      const b = document.createElement('button');
      b.id = 'b8s-fatal';
      b.textContent = '⚠ Tool 800BEST lỗi: ' + String(e && e.message || e).slice(0, 120) + ' (bấm để sao chép)';
      b.style.cssText = 'position:fixed;right:20px;bottom:22px;z-index:2147483647;background:#b91c1c;color:#fff;border:0;border-radius:10px;padding:10px 14px;font:600 13px sans-serif;cursor:pointer;max-width:70vw';
      b.onclick = () => { try { navigator.clipboard.writeText(String(e && e.stack || e)); } catch { /* bỏ qua */ } };
      document.documentElement.appendChild(b);
    } catch { /* bỏ qua */ }
  }
  const safeMount = () => { try { mount(); } catch (e) { showFatal(e); } };
  try { hookNetwork(); } catch (e) { console.error('[800BEST] hookNetwork', e); }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', safeMount, { once: true });
  else safeMount();
  setInterval(() => {
    if (document.body && !document.getElementById('b8s-launch') && !document.getElementById('b8s-fatal')) {
      document.getElementById('b8s-app')?.remove();
      document.getElementById('b8s-style')?.remove();
      safeMount();
    }
  }, 2500);
})();
