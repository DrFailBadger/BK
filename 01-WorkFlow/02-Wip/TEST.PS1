$CamlQueryNotAssignedtoUserCompleted =
"<View>
    <Query>  
    <OrderBy>
    <FieldRef Name='ID' />
    </OrderBy>
    <Where>
    <Neq>
    <FieldRef Name='Status' />
    <Value Type='Choice'>Completed</Value>
    </Neq>
    </Where>
    </Query>
</View>"
$CamlQueryNotAssignedtoUserCompleted