t = int(input())
for test_case in range(t):
    n = int(input())
    int_list = list(map(int, input().split()))

    first_to_last = [int_list[0]]
    sum_first_to_last = int_list[0]
    last_to_first = [int_list[n-1]]
    sum_last_to_fast = int_list[n-1]
    i = 1
    j = n - 2
    while i < n-1:
        first_to_last.append(sum_first_to_last + int_list[i])
        sum_first_to_last += int_list[i]
        last_to_first.append(sum_last_to_fast + int_list[j])
        sum_last_to_fast += int_list[j]
        i += 1
        j -= 1
    # print(first_to_last, last_to_first)

    f_to_l_ind = 0
    l_to_f_ind = 0

    ok_sum = 0
    stop = 0
    for i in range(n-1):
        if not stop:
            for j in range(l_to_f_ind, n-1):
                # print(i, l_to_f_ind)
                if first_to_last[i] == last_to_first[j]:
                    if (i+j+2) <= n:
                        ok_sum = i + j + 2
                        l_to_f_ind = j
                        break
                    else:
                        stop = 1
                        break
                elif first_to_last[i] < last_to_first[j]:
                    l_to_f_ind = j
                    if (i+j+2) >= n:
                        stop = 1
                        break
                    else:
                        break
        else:
            break
    print(ok_sum)
