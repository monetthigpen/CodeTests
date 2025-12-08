<TagPickerList className="tagpickerList">
  {React.Children.map(children, (child) => {
    if (!React.isValidElement(child)) return child;

    // get the value prop (DisplayText) from the TagPickerOption
    const val = (child.props as any).value as string | undefined;
    if (!val || val === "no-matches") return child;

    // find the matching entity so we can grab the Title (position)
    const ent = optionRaw.find(
      (v) => v.DisplayText.toLowerCase() === val.toLowerCase()
    );
    const role = ent?.EntityData?.Title ?? "";

    // add the secondary line (position) and keep everything else the same
    return React.cloneElement(child, {
      secondaryContent: role,
    });
  })}
</TagPickerList>








