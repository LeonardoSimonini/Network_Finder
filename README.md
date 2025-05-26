# FIND_NETWORK
FIND_NETWORK is a function written in VBA language to make it usable in Microsoft Excel

## Why this function
Recently I worked on a Microsoft Excel Sheet where I had a long list of IP address and a list of Network, my aim was to associate each IP address to a Network in the list. No need to say that it could have take a lot of time doing it one by one manually. So I searched online a function that could have helped me in this task, a custom function to use in Microsoft Excel. Maybe I have not searched enough, but I didn't find it so I create that funcion myself. Than I thought that someone else could have been in a similar situation so I decided to publish that function and here we are. FIND_NETWORK is the function I created, so you can too automate this kind of task with Microsoft Excel.

## How to use it
FIND_NETWORK take two argument in input:
  1. The first one is an IP address, which I need to associate with a Network.
  2. The second one is a range of cell in Excel with a list of Networks.

The **output** of the function is the network which owns the IP address. If the right network is not in the list, the function return an error. If in the list there are network and subnetworks which the IP could belongs to, the function returns the smallest subnetwork the IP address belongs to.

![image](https://github.com/user-attachments/assets/5718ca67-96a9-4f99-930a-7f9b87c720da)
