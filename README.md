# PowerPointDarkThemeFix

Default PowerPoint's color picker messes grey colors for fill, borders, shadow and glow. PowerPointDarkThemeFix does this simple task right.

![preview](https://user-images.githubusercontent.com/34414488/127713963-fd0b2e00-a898-4dc9-b4ff-210e53e69aa2.png)

This Add-In adds a new group "Dark Theme Fix" to the "Home" and "Drawing Tools" tabs. This group contains clones of the default PowerPoint color picker and other controls.

A new color picker applies right colors, though, it provides only default color scheme. To use a color scheme provided by the theme, use default controls.

There's also a "Fix All" button which will fix all opened presentations.

Though, [you might not need this Add-In](#alternative-approach-to-the-problem).

If you find this Add-In useful, you also might want to get [ExcelDarkThemeFix](https://github.com/matafokka/ExcelDarkThemeFix).

# Installation

## Versions overview

There're four versions of this Add-In, pick one you like the most. If you're not sure, pick **3_columns**, it's the most comfortable, but uses the most horizontal space out of all.

**3_columns** - the way it was intended; there are three columns, every menu has its own button.

![3 cols](https://user-images.githubusercontent.com/34414488/127712015-765a4b80-64c2-4a74-9703-f339294b693a.png)

**2_columns_no_duplicates** - two columns; menus with duplicate functionality has been removed, though, they still can be found in the "Drawing" panel.

![2 no defs](https://user-images.githubusercontent.com/34414488/127712013-36589127-b07e-4900-9650-c71075178b37.png)

**two_columns_merged** - two columns; dash options has been merged into "Dashes" menu; shadow and glow color menus has been moved to "Shape Effects" menu.

![2 merged](https://user-images.githubusercontent.com/34414488/127712011-c6c05d23-0d0a-46e0-8a46-cfc098b2c160.png)

**1_column** - one column; additional fill and dashes menus has been removed; shadow and glow color menus has been moved to "Shape Effects" menu; "Fix All" button has been moved to all three menus.

![1 col](https://user-images.githubusercontent.com/34414488/127712006-50d0838a-0006-4b79-8557-8211c04d47aa.png)

## Installation process

Basically, you just need to install an Add-In. You can either stick to the instructions below (copied from MS website) or find a better guide.

1. Download [the latest release](https://github.com/matafokka/PowerPointDarkThemeFix/releases/latest). Don't download files by clicking the big green button that says "Code"! It's sources, you don't need them. Download from the provided link.
1. "Unblock" the downloaded file:
    1. Move the cursor over the file, click right mouse button and click "Properties".
    1. At the bottom of the window find "Unblock" checkbox and check it. Then click "OK".
    1. If you don't have this checkbox, please, proceed to the rest of the steps.
1. Move the downloaded file where it won't get in your way. I'd suggest `%AppData%\Microsoft\AddIns`.
1. Open PowerPoint.
1. Click the "File" tab, and then click "Options".
1. In the "Options" dialog box, click "Add-Ins".
1. In the "Manage" list at the bottom of the dialog box, click "PowerPoint Add-ins", and then click "Go".
1. In the "Add-Ins" dialog box, click "Add New".
1. In the "Add New PowerPoint Add-In" dialog box, open a downloaded file, and then click "OK".
1. A security notice might appear. Click "Enable Macros", and then click "Close".

# Compatibility list

Fully compatible with Microsoft Office 2016 and 2019. Should also work with 2010 and 2013, but I've not tested. Don't know about 365.

# Alternative approach to the problem

The problem affects only grey colors (including white and black). The solution is not to use these colors not only because of this problem but because these colors will make your presentation look muddy.

This fact, or rather rule, comes from digital art, and the reason for this is that monitor works differently from paper or any other real life material.

Instead, use shades of grey (or white, or black) with a bit of color added to them. Choose a color that matches your presentation's style.

If you don't have a defined style, pick bluish colors for shadow and yellowish for lights.

Though, there are cases when you'll want to use straight grey colors, such as black and white style, so this Add-In still nice to have.

# FAQ

## Why new color picker is different from the default? Also, why not change existing one?

Because you can't modify the complex default controls such as color picker. Moreover, all color pickers are used only for a specific purpose, there're separate color pickers for different uses. You also can't add submenus to the color picker, it can contain only buttons, and there's no alternative.

## Why there's no color preview on mouse over?

Because there's no such event. This feature is impossible to implement.

## Why there're both new and default controls?

As said earlier, a new color picker provides only default color scheme. Even if it's possible, it would be hard to implement such a feature. So, you still might need default color pickers.

## I have a problem or found a bug. Where can I report it?

Please, fill out an issue on this repository and provide as much information as possible.

If a problem is not related to the grey colors, please, report it too. I'll do my best to fix PowerPoint for you.

## How can I contribute to the project?

If you're regular user, you can test this Add-In with different kinds of presentations and report found issues, or install an older version of Microsoft Office and check if this Add-In works with it.

You can also spread the word about this Add-In (especially in non-English speaking parts of the Internet), so other people could fix their PowerPoint.
