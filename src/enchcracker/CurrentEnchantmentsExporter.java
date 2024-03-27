package enchcracker;

import org.apache.poi.ss.usermodel.Cell; 
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.xssf.usermodel.XSSFRow; 
import org.apache.poi.xssf.usermodel.XSSFSheet; 
import org.apache.poi.xssf.usermodel.XSSFWorkbook; 

import java.util.Random;
import java.util.List;
import java.util.Map; 
import java.util.Set; 
import java.util.TreeMap;
import java.io.IOException;
import java.io.File;
import java.io.FileOutputStream; 

public class CurrentEnchantmentsExporter {
    // referenced from app
	long playerSeed;
    Enchantments myEnchantments;
    boolean[][] itemIdCanHaveMatEnchTier = new boolean[35][8];

    // recalc from app
    int[][] enchantLevels = new int[16][3];

    // mine
    Items[] armors = { Items.HELMET, Items.CHESTPLATE, Items.LEGGINGS, Items.BOOTS, Items.WOLF_ARMOR};
    Items[] shields = { Items.SPARTAN_SHIELD, Items.BS_SHIELD};
    Items[] tools = { Items.PICKAXE, Items.SHOVEL_MATTOCK_SAW, Items.AXE};
    Items[] melees = { Items.WOLF_ARMOR, Items.AXE, Items.BS_DAGGER, Items.KNIFE, Items.WEAPON_WITHOUT, Items.WEAPON_WITH, Items.BS_BATTLE_AXE, Items.SPARTAN_BATTLEAXE};
    Items[] ranged = { Items.SWITCH_BOW, Items.SWITCH_CROSSBOW, Items.BOW, Items.CROSSBOW};
    Items[] books = { Items.BOOK };
    // lol
    // 	FISHING_ROD("Fishing Rod",21),
    // 	THROWING_WEAPON("Throwing Weapon",25),
    Short[] tier_colors = {
            IndexedColors.TAN.getIndex(), 
            IndexedColors.GREY_50_PERCENT.getIndex(), 
            IndexedColors.LIGHT_BLUE.getIndex(), 
            IndexedColors.BROWN.getIndex(), 
            IndexedColors.VIOLET.getIndex(), 
            IndexedColors.YELLOW.getIndex(), 
            IndexedColors.GOLD.getIndex(), 
            IndexedColors.BLACK.getIndex()
        };

    // Nisch work stealer, info to reset local rand
    public CurrentEnchantmentsExporter(long playerSeed, Enchantments myEnchantments, boolean[][] itemIdCanHaveMatEnchTier) {
		super();

		this.playerSeed = playerSeed;
        this.myEnchantments = myEnchantments;
        try {
            this.itemIdCanHaveMatEnchTier = itemIdCanHaveMatEnchTier;
        } catch (ArrayIndexOutOfBoundsException e) {
            System.out.println("Error: Index is out of bounds.");
            e.printStackTrace();
        }

        // imagine messing up main app's rand
		Random rand = new Random();
        int xpSeed = (int) (playerSeed >>> 16);

        // copied from main app, original calcs trapped in an ActionListener lololol
        for (int bookshelves = 0; bookshelves <= 15; bookshelves++) {
            rand.setSeed(xpSeed); // I forgot this once
            for (int slot = 0; slot < 3; slot++) {
                int level = myEnchantments.calcEnchantmentTableLevel(rand, slot, bookshelves, 1);
                if (level < slot + 1) {
                    level = 0;
                }
                enchantLevels[bookshelves][slot] = level;
            }
        }
	}

    // export to xlsx
    public void export() {
		// workbook object 
        XSSFWorkbook workbook = new XSSFWorkbook(); 
  
        // spreadsheet objects
        XSSFSheet booksheet = workbook.createSheet("Book Enchants"); 
        XSSFSheet armorsheet = workbook.createSheet("Armor Enchants"); 
        XSSFSheet meleesheet = workbook.createSheet("Melee Enchants"); 
        XSSFSheet shieldsheet = workbook.createSheet("Shield Enchants"); 
        XSSFSheet toolsheet = workbook.createSheet("Tool Enchants"); 
        XSSFSheet rangedsheet = workbook.createSheet("Ranged Enchants"); 

        buildItemSheet(booksheet, books);
        buildItemSheet(armorsheet, armors);
        buildItemSheet(meleesheet, melees);
        buildItemSheet(shieldsheet, shields);
        buildItemSheet(toolsheet, tools);
        buildItemSheet(rangedsheet, ranged);

		try {
			// .xlsx is the format for Excel Sheets... 
			// writing the workbook into the file... 
			FileOutputStream out = new FileOutputStream(new File("Current_Enchants [" + playerSeed + "].xlsx")); 
			workbook.write(out); 
        	out.close(); 
        } catch(IOException e) {
            System.err.println("Failed to open file.");
            e.printStackTrace();
        }

	}

    // helper method, adds enchants and format strings to map and return highest number of enchants
    private int enchantsToMap(Map<Integer, Object[]> enchantsList, int row_id, int item_id) {
        int highest_num_enchantments = 0;

        Random rand = new Random();
        int xpSeed = (int) (playerSeed >>> 16);
        rand.setSeed(xpSeed);

        for (int current_tier = 0; current_tier < itemIdCanHaveMatEnchTier[0].length; current_tier++) {
            if (itemIdCanHaveMatEnchTier[item_id][current_tier]) {
                // approximate enchantability, it works I promise
                // would need Materials.java for exact numbers
                int currentRoundedEnchantabilities = (current_tier * 4 + 1);
                enchantsList.put(
                        row_id++, 
                        new Object[] { tier_colors[current_tier] }
                    ); 
                enchantsList.put( 
                    row_id++, 
                    new Object[] { "Tier: " + current_tier}
                    ); 

                for (int bookshelves = 0; bookshelves <= 15; bookshelves++) {
                    enchantsList.put(
                        row_id++, 
                        new Object[] {"Bookshelves: " + bookshelves}
                    ); 
                    for (int slot = 0; slot < 3; slot++) {
                        List<Enchantments.EnchantmentInstance> enchantments 
                            = myEnchantments.getEnchantmentsInTable(rand, xpSeed, item_id, currentRoundedEnchantabilities, slot, enchantLevels[bookshelves][slot]);
                        highest_num_enchantments = (highest_num_enchantments < enchantments.size()) ? enchantments.size() : highest_num_enchantments;
                        enchantsList.put(
                            row_id++, 
                            new Object[] {"Levels: " + enchantLevels[bookshelves][slot]}
                        ); 
                        enchantsList.put(
                            row_id++, 
                            enchantments.toString().substring(1, enchantments.toString().length()-1).split(",")
                        ); 
                    }
                    enchantsList.put(
                        row_id++, 
                        new Object[] { tier_colors[current_tier] }
                    ); 
                }

                enchantsList.put(
                    row_id++, 
                    new Object[] {""}
                );
            }
        }


        return highest_num_enchantments;
    }

    // create sheet based on array of items
    private void buildItemSheet(XSSFSheet spreadsheet, Items[] items) {
        // creating a row object 
        XSSFRow row;

        // formatting
        CellStyle textStyle = spreadsheet.getWorkbook().createCellStyle();
        int displacement = 0;

  
        for (int i = 0; i < items.length; i++) {
            // This data needs to be written (Object[]) 
            Map<Integer, Object[]> enchantsList = new TreeMap<Integer, Object[]>(); 
            int row_id = 1; // confusing name, but key to the map
            
            String item_name = items[i].name;
            int item_id = items[i].id;

            enchantsList.put( 
                row_id++, 
                new Object[] { "Item: " + item_name}
                ); 

            int highest_num_enchantments = enchantsToMap(enchantsList, row_id, item_id);
    
            Set<Integer> keyid = enchantsList.keySet(); 
            int rowid = 0; // now you see why it was a bad name
    
            // writing the data into the sheets... 
            // item types in columns
            // enchantibility tiers in vertical groups
            for (int key : keyid) { 
                row = spreadsheet.getRow(rowid); 
                if (row == null)
                    row = spreadsheet.createRow(rowid); 
                rowid++;

                Object[] objectArr = enchantsList.get(key); 
                int cellid = displacement;
    
                for (Object obj : objectArr) {
                    Cell cell = row.createCell(cellid++); 

                    if (obj instanceof String) {
                        cell.setCellValue((String)obj); 
                    }
                    // what could go wrong?
                    else if (obj instanceof Short) {
                        CellStyle borderStyle = spreadsheet.getWorkbook().createCellStyle();
                        borderStyle.setFillForegroundColor((Short)obj);
                        borderStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                        cell.setCellStyle(borderStyle);
                    }
                } 
            } 
            displacement += highest_num_enchantments + 1;
        }
        for (int i = 0; i < displacement; i++) {
            spreadsheet.autoSizeColumn(i);
        }
        spreadsheet.createFreezePane(0, 1);

    }


}