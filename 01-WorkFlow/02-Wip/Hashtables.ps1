
$HT = @{
    Key  = 'Value';
    'Associate this' = 'with that'; 
    word = 'definiton'
}

$hash = @{ ID = 1; Shape = "Square"; Color = "Blue"}

$ht | gm
$ht.Key
$ht.Keys #defaults
$ht.Values #defaults
$ht.word
$ht.'Associate this'