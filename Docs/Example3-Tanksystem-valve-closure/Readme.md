# Simple tanksystem 2 - multiple flowpaths / branches

## Pipe layout
![alt text](Tanksystem2.png "Logo Title Text 1")

## Step 0: Build on tanksystem 1

## Step 1: Create a new variable trip

The input for a variable trip that is true >= 15 seconds looks like this

`405  time  0   ge  null  0   10.00   n`

To include this trip one can either use the `Custom`-word and include it as a custom include file or it can be added directly like this:

![alt text](Add-trip.gif "Logo Title Text 1")

## Step 2: Create a new include file for valve V1
![alt text](Create-includefile.gif "Logo Title Text 1")


## Step 3: Generate input file and review output

## Step 4: Run calculation and look at result
