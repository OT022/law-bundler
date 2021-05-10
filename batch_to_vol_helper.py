def split_to_volumes(bundle_content: list, max_volume_page_count):

    volumes = [[]]
    bundle_page_count = 0
    volume_index = 0
    volume_page_count = 0
    
    for item in bundle_content:
        #skip first row which contains headers
        if len(item) != 9:
            print("row length wrong - ", item)
            continue

        item_page_count = int(item[8])

        if (volume_page_count + item_page_count) <= max_volume_page_count:
            volumes[volume_index].append(item)
            volume_page_count = volume_page_count + item_page_count
        else:
            volumes.append([])
            volume_index += 1
            volumes[volume_index].append(item)
            volume_page_count = item_page_count
        
        bundle_page_count += item_page_count

    print("bundle_page_count",bundle_page_count)

    return volumes
