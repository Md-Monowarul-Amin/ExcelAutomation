t = int(input())
for test_case in range(t):
    n, k = map(int, input().split())
    str_ = input()

    ans = k
    w_in_str = 0
    for i in range(k):
        if str_[i] == "W":
            w_in_str += 1
    i = 0
    j = k
    ans = min(w_in_str, ans)
    for i in range(n-k):
        if str_[i] == "W" and str_[j] == "B":
            w_in_str -= 1
        elif str_[i] == "B" and str_[j] == "W":
            w_in_str += 1
        j += 1
        ans = min(w_in_str, ans)

    print(ans)
